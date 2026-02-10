// [2025-11-29] v2.11 - Fix Date Highlighting Mismatch
// Changes:
// - Added 'areValuesEquivalent' helper to handle Excel Serial Numbers vs Date Strings.
// - Comparison logic now normalizes dates to 'MM/DD/YY' before flagging changes.
// - Prevents false positive highlighting when formats differ (e.g. "11/25/2025" vs "11/25/25").

import React, { useEffect } from 'react';
import {
	normalizeKey,
	parseHyperlink,
	makeHyperlinkFormula,
	findColumnIndex,
	buildAliasMap,
	renameHeaderArray,
	gatherIdentifierColumn,
	renameObjectRows,
	renameArrayRows,
	computeSavedStaticFromValues,
	computeIdentifierListFromValues,
	applyStaticColumnsWithContext,
	extractHyperlink,
	getHeaderIndexMap,
} from './dataProcessorUtility';

export default function DataProcessor({ data, sheetName, refreshSheetName, settingsColumns, matched, onComplete, onStatus, action, conditionalFormat }) {
	const notifyComplete = (payload) => {
		if (typeof onComplete === 'function') {
			try { onComplete(payload); } catch (e) { /* swallow callback errors */ }
		}
	};

	const notifyStatus = (msg) => {
		if (typeof onStatus === 'function') {
			try { onStatus(String(msg)); } catch (e) { /* swallow callback errors */ }
		}
	};

    // --- NEW: Comparison Helpers ---

    // Convert Excel Serial Number (e.g. 45615) to JS Date
    const excelSerialToDate = (serial) => {
        // Excel base date is Dec 30, 1899
        const utc_days  = Math.floor(serial - 25569);
        const utc_value = utc_days * 86400;                                        
        const date_info = new Date(utc_value * 1000);
        return date_info;
    };

    // Normalize values for comparison (Handles Dates, Numbers vs Strings, etc.)
    const areValuesEquivalent = (val1, val2) => {
        // 1. Handle Nulls/Undefined
        const v1 = val1 == null ? '' : val1;
        const v2 = val2 == null ? '' : val2;

        if (v1 === v2) return true;

        // 2. Loose Equality (for 5 vs "5")
        // eslint-disable-next-line eqeqeq
        if (v1 == v2) return true;

        // 3. Date Intelligence
        // We want to normalize everything to MM/DD/YY for comparison if it looks like a date.
        
        const toDateString = (v) => {
            if (typeof v === 'number' && v > 20000 && v < 60000) {
                // Likely an Excel Serial Number
                const d = excelSerialToDate(v);
                if (!isNaN(d.getTime())) {
                    const mm = String(d.getUTCMonth() + 1).padStart(2, '0');
                    const dd = String(d.getUTCDate()).padStart(2, '0');
                    const yy = String(d.getUTCFullYear()).slice(-2);
                    return `${mm}/${dd}/${yy}`;
                }
            }
            
            if (typeof v === 'string') {
                // Check if it looks like a date (MM/DD/YY or YYYY-MM-DD)
                if (v.match(/^\d{1,2}[\/-]\d{1,2}[\/-]\d{2,4}$/) || v.match(/^\d{4}[\/-]\d{1,2}[\/-]\d{1,2}$/)) {
                   const d = new Date(v);
                   if (!isNaN(d.getTime())) {
                        // Use UTC methods to align with ImportManager's logic
                        const mm = String(d.getUTCMonth() + 1).padStart(2, '0');
                        const dd = String(d.getUTCDate()).padStart(2, '0');
                        const yy = String(d.getUTCFullYear()).slice(-2);
                        return `${mm}/${dd}/${yy}`;
                   }
                }
            }
            return null; // Not a date
        };

        const d1 = toDateString(v1);
        const d2 = toDateString(v2);

        if (d1 && d2) {
            return d1 === d2;
        }

        // 4. String trim fallback
        return String(v1).trim() === String(v2).trim();
    };

	// Helper to queue formatting commands for a list of addresses.
	function queueFormatRanges(sheet, addresses, color) {
		if (!addresses || addresses.length === 0) return;
		
        const BATCH_SIZE = 1000;
        for (let i = 0; i < addresses.length; i += BATCH_SIZE) {
            const batch = addresses.slice(i, i + BATCH_SIZE);
            batch.forEach(addr => {
                try {
                    const range = sheet.getRange(addr);
                    if (color) {
                        range.format.fill.color = color;
                    } else {
                        range.format.fill.clear();
                    }
                } catch (e) {
                    console.warn('Skipping invalid range address:', addr);
                }
            });
        }
	}

	// Helper: Apply Conditional Formatting
	function applyConditionalFormatting(sheet, columnIndex, formatRule) {
		console.log(`[DataProcessor] Applying Conditional Format: Rule=`, formatRule, ` Index=${columnIndex}`);
		
		if (!formatRule || !formatRule.column || columnIndex < 0) {
			console.warn('[DataProcessor] Skipped CF: Invalid rule or index');
			return;
		}

		try {
			const range = sheet.getRangeByIndexes(0, columnIndex, 1, 1).getEntireColumn();
			range.conditionalFormats.clearAll();

			if (formatRule.condition === 'Color Scales' && formatRule.format === 'G-Y-R Color Scale') {
				const cf = range.conditionalFormats.add(Excel.ConditionalFormatType.colorScale);
				cf.colorScale.criteria = {
					minimum: { formula: null, type: Excel.ConditionalFormatColorCriterionType.lowestValue, color: "#F8696B" },
					midpoint: { formula: null, type: Excel.ConditionalFormatColorCriterionType.percentile, value: 50, color: "#FFEB84" },
					maximum: { formula: null, type: Excel.ConditionalFormatColorCriterionType.highestValue, color: "#63BE7B" }
				};
				if (formatRule.column.toLowerCase().includes('missing')) {
					cf.colorScale.criteria = {
						minimum: { formula: null, type: Excel.ConditionalFormatColorCriterionType.lowestValue, color: "#63BE7B" },
						midpoint: { formula: null, type: Excel.ConditionalFormatColorCriterionType.percentile, value: 50, color: "#FFEB84" },
						maximum: { formula: null, type: Excel.ConditionalFormatColorCriterionType.highestValue, color: "#F8696B" }
					};
				}
			} else if (formatRule.condition === 'Highlight Cells with' && Array.isArray(formatRule.format)) {
				const type = formatRule.format[0];
				const operator = formatRule.format[1];
				const val = formatRule.format[2];
				const style = formatRule.format[3];

				if (type === 'Specific text' && operator === 'Beginning with') {
					const cf = range.conditionalFormats.add(Excel.ConditionalFormatType.containsText);
					cf.textComparison.format.font.color = "#006100"; 
					cf.textComparison.format.fill.color = "#C6EFCE"; 

					if (style.includes('Red')) {
						cf.textComparison.format.font.color = "#9C0006";
						cf.textComparison.format.fill.color = "#FFC7CE";
					} else if (style.includes('Yellow')) {
						cf.textComparison.format.font.color = "#9C6500";
						cf.textComparison.format.fill.color = "#FFEB9C";
					}
					cf.textComparison.rule = { operator: Excel.ConditionalTextOperator.beginsWith, text: val };
				}
			}
		} catch (err) {
			console.error('[DataProcessor] Error queuing conditional formatting', err);
		}
	}


	// Refresh: write `data` to the target worksheet.
	async function Refresh(dataToWrite, targetSheetName = sheetName, suppressCompletion = false) {
		notifyStatus(`Starting refresh on ${targetSheetName}...`);
		if (!dataToWrite || !dataToWrite.length) {
			notifyStatus('No data to write.');
			const result = { success: false, reason: 'no-data' };
			if (!suppressCompletion) notifyComplete(result);
			return result;
		}
		if (!targetSheetName) {
			notifyStatus('Missing sheet name.');
			const result = { success: false, reason: 'missing-sheetName' };
			if (!suppressCompletion) notifyComplete(result);
			return result;
		}
		if (!window.Excel || !window.Excel.run) {
			notifyStatus('Excel JS API not available.');
			const result = { success: false, reason: 'no-excel-api' };
			if (!suppressCompletion) notifyComplete(result);
			return result;
		}

		let rowCount = 0;
		try {
			notifyStatus('Reading existing worksheet...');
			
			// 1. READ PHASE
			let usedValues = [];
			let usedFormulas = [];
			let usedProps = [];
			let sheetExisted = false;

			await Excel.run(async (context) => {
				const sheets = context.workbook.worksheets;
				const sheet = sheets.getItemOrNullObject(targetSheetName);
				await context.sync();
				if (!sheet.isNullObject) {
					sheetExisted = true;
					const used = sheet.getUsedRangeOrNullObject();
					used.load(['values', 'formulas']);
					const propsResult = used.getCellProperties({ format: { fill: { color: true } } });
					await context.sync();
					if (!used.isNullObject) {
						usedValues = used.values || [];
						usedFormulas = used.formulas || [];
						usedProps = propsResult.value || [];
					}
				}
			});

			notifyStatus('Processing data structure...');

			// derived headers
			let sheetHeaders = (sheetExisted && usedValues.length > 0)
				? usedValues[0].map(v => (v == null ? '' : String(v).trim()))
				: [];

			const dataFirst = dataToWrite[0];
			const isObjectRows = (dataFirst && typeof dataFirst === 'object' && !Array.isArray(dataFirst));
			let dataHeaders = isObjectRows 
				? Object.keys(dataFirst).map(h => (h == null ? '' : String(h).trim()))
				: (Array.isArray(dataFirst) ? dataFirst.map(h => (h == null ? '' : String(h).trim())) : []);
			
			const dataHeaderIndexMap = getHeaderIndexMap(dataHeaders);

			if (sheetHeaders.length === 0 && dataHeaders.length > 0) {
				sheetHeaders = [...dataHeaders];
			}
			
			const hasHeaders = sheetHeaders.some(h => h !== '');
			const sheetHeaderIndexMap = getHeaderIndexMap(sheetHeaders);

			const settingsCols = Array.isArray(settingsColumns) ? settingsColumns : [];
			const identifierSetting = gatherIdentifierColumn(settingsCols);

			// Capture PRE-EXISTING identifiers AND Data
			const beforeIdentifiers = new Set();
            const beforeDataMap = new Map(); // Map<ID, RowArray>

			if (identifierSetting && usedValues.length > 0) {
				const beforeInfo = computeIdentifierListFromValues(usedValues, sheetHeaders, identifierSetting, usedFormulas);
                if (beforeInfo.identifierIndex !== -1) {
                    beforeInfo.identifiers.forEach(id => beforeIdentifiers.add(id));
                    for (let r = 1; r < usedValues.length; r++) {
                        const row = usedValues[r];
                        let rawId = row[beforeInfo.identifierIndex];
                        if (usedFormulas[r] && typeof usedFormulas[r][beforeInfo.identifierIndex] === 'string' && usedFormulas[r][beforeInfo.identifierIndex].startsWith('=')) {
                            const parsed = parseHyperlink(usedFormulas[r][beforeInfo.identifierIndex]);
                            if (parsed.url) rawId = parsed.url;
                        }
                        const idStr = rawId == null ? '' : String(rawId).trim();
                        if (idStr) {
                            beforeDataMap.set(idStr, row);
                        }
                    }
                }
			}

			// Normalization & Aliasing Logic
			const matchedByName = new Map();
			for (const sc of settingsCols) {
				if (!sc || !sc.name) continue;
				const nameKey = normalizeKey(sc.name);
				if (sheetHeaderIndexMap.has(nameKey) && dataHeaderIndexMap.has(nameKey)) {
					matchedByName.set(nameKey, String(sc.name).trim());
				}
			}

			const remainingSettings = settingsCols.filter(sc => sc && sc.name && !matchedByName.has(normalizeKey(sc.name)));
			const sheetAliasMap = buildAliasMap(remainingSettings, sheetHeaders);
			const dataAliasMap = buildAliasMap(remainingSettings, dataHeaders);

			if (matchedByName.size > 0) {
				sheetHeaders = sheetHeaders.map(h => {
					const key = normalizeKey(h);
					return matchedByName.has(key) ? matchedByName.get(key) : h;
				});
				Object.assign(sheetHeaderIndexMap, getHeaderIndexMap(sheetHeaders));
			}
			if (sheetAliasMap.size > 0) {
				sheetHeaders = renameHeaderArray(sheetHeaders, sheetAliasMap);
			}

			const combinedDataAliasMap = new Map([...matchedByName, ...dataAliasMap]);
			if (combinedDataAliasMap.size > 0) {
				const firstRow = dataToWrite[0];
				if (firstRow && typeof firstRow === 'object' && !Array.isArray(firstRow)) {
					dataToWrite = renameObjectRows(dataToWrite, combinedDataAliasMap);
				} else if (Array.isArray(firstRow)) {
					dataToWrite = renameArrayRows(dataToWrite, combinedDataAliasMap);
				}
			}

			// --- Merge new columns from imported data into sheet headers ---
			let newColumnsFromImport = [];
			if (sheetExisted && sheetHeaders.length > 0) {
				const renamedFirst = dataToWrite[0];
				let renamedDataHeaders = [];
				if (renamedFirst && typeof renamedFirst === 'object' && !Array.isArray(renamedFirst)) {
					renamedDataHeaders = Object.keys(renamedFirst).map(h => String(h).trim());
				} else if (Array.isArray(renamedFirst)) {
					renamedDataHeaders = renamedFirst.map(h => (h == null ? '' : String(h).trim()));
				}

				const sheetNormSet = new Set(sheetHeaders.map(h => normalizeKey(h)));
				for (const dh of renamedDataHeaders) {
					const norm = normalizeKey(dh);
					if (norm && !sheetNormSet.has(norm)) {
						sheetHeaders.push(dh);
						sheetNormSet.add(norm);
						newColumnsFromImport.push(dh);
					}
				}
				if (newColumnsFromImport.length > 0) {
					notifyStatus(`Adding ${newColumnsFromImport.length} new column(s) from import: ${newColumnsFromImport.join(', ')}`);
				}
			}

			let dataFieldSet = new Set();
			const firstAfter = dataToWrite[0];
			if (firstAfter && typeof firstAfter === 'object' && !Array.isArray(firstAfter)) {
				dataToWrite.forEach(row => Object.keys(row || {}).forEach(k => dataFieldSet.add(normalizeKey(k))));
			} else if (Array.isArray(firstAfter)) {
				firstAfter.forEach(v => dataFieldSet.add(normalizeKey(v)));
			}

            const columnsInImport = new Set();
            sheetHeaders.forEach((h, i) => {
                const key = normalizeKey(h);
                if (key && dataFieldSet.has(key)) {
                    columnsInImport.add(i);
                }
            });

			const staticCols = [];
			sheetHeaders.forEach(h => {
				const key = normalizeKey(h);
				if (key && !dataFieldSet.has(key)) staticCols.push(h);
			});

			let savedStatic = null;
			if (staticCols.length > 0 && identifierSetting) {
				savedStatic = computeSavedStaticFromValues(usedValues, sheetHeaders, staticCols, identifierSetting, usedFormulas, usedProps);
			}

			// Prepare Matrix
			notifyStatus('Constructing write matrix...');
			
			let rows;
			let writeStartRow = 0;

			if (hasHeaders) {
				const desiredColumns = sheetHeaders.map((h, idx) => (h ? { name: h, index: idx } : null)).filter(Boolean);
				rows = new Array(dataToWrite.length);
				for (let i = 0; i < dataToWrite.length; i++) {
					const item = dataToWrite[i];
					const rowArr = new Array(desiredColumns.length);
					for (let j = 0; j < desiredColumns.length; j++) {
						const col = desiredColumns[j];
						if (isObjectRows) rowArr[j] = (item && Object.prototype.hasOwnProperty.call(item, col.name)) ? item[col.name] : '';
						else rowArr[j] = (Array.isArray(item) && item.length > col.index) ? item[col.index] : '';
					}
					rows[i] = rowArr;
				}
				writeStartRow = 1;
			} else {
				if (isObjectRows) {
					const headers = Object.keys(firstAfter || {});
					rows = new Array(dataToWrite.length + 1);
					rows[0] = headers;
					for (let i = 0; i < dataToWrite.length; i++) {
						rows[i + 1] = headers.map(h => dataToWrite[i][h] ?? '');
					}
					writeStartRow = 0;
				} else {
					rows = dataToWrite.slice();
					writeStartRow = 0;
				}
			}

			rowCount = rows.length;
			let colCount = rows[0] ? rows[0].length : 0;
			if (rowCount === 0) {
				const result = { success: false, reason: 'nothing-to-write' };
				if (!suppressCompletion) notifyComplete(result);
				return result;
			}

			// Calculate Highlights
			const rangesToHighlight = [];
            const cellsToHighlight = [];
			const rangesToClear = [];
			
			if (identifierSetting) {
				const idIdx = getHeaderIndexMap(sheetHeaders).get(normalizeKey(identifierSetting.name));
				
				if (idIdx !== undefined && idIdx >= 0) {
					const getColLetter = (c) => {
						let letter = '';
						while (c >= 0) {
							letter = String.fromCharCode((c % 26) + 65) + letter;
							c = Math.floor(c / 26) - 1;
						}
						return letter;
					};
					const startColLetter = getColLetter(0); 
					const endColLetter = getColLetter(Math.max(colCount, 1) - 1);

					let currentBlockType = null; 
					let currentBlockStart = -1;
					
					const closeBlock = (endRowIndex) => {
						if (currentBlockType && currentBlockStart >= 0) {
							const r1 = currentBlockStart + 1;
							const r2 = endRowIndex + 1;
							const address = `${startColLetter}${r1}:${endColLetter}${r2}`;
							
							if (currentBlockType === 'highlight') rangesToHighlight.push(address);
							else if (currentBlockType === 'clear') rangesToClear.push(address);
						}
						currentBlockType = null;
						currentBlockStart = -1;
					};

					for (let r = 0; r < rows.length; r++) {
						const rowData = rows[r];
						let val = rowData[idIdx]; 
						if (typeof val === 'string' && val.startsWith('=')) {
							const link = extractHyperlink(val);
							if (link) val = link;
						}
						const id = val == null ? '' : String(val).trim();
						
						if (!id) {
							closeBlock(writeStartRow + r - 1);
							continue;
						}

						const isExisting = beforeIdentifiers.has(id);
						const type = !isExisting ? 'highlight' : 'clear';
						const absRowIndex = writeStartRow + r;

						if (type !== currentBlockType) {
							closeBlock(absRowIndex - 1);
							currentBlockType = type;
							currentBlockStart = absRowIndex;
						}

                        // Check specific cells in existing rows
                        if (isExisting) {
                            const oldRow = beforeDataMap.get(id);
                            if (oldRow) {
                                for (let c = 0; c < rowData.length; c++) {
                                    if (!columnsInImport.has(c)) continue;

                                    const newVal = rowData[c];
                                    let oldVal = '';
                                    if (c < oldRow.length) oldVal = oldRow[c];

                                    // V2.11 Fix: Use smart comparison
                                    if (!areValuesEquivalent(newVal, oldVal)) {
                                        const cellAddr = `${getColLetter(c)}${absRowIndex + 1}`;
                                        cellsToHighlight.push(cellAddr);
                                    }
                                }
                            }
                        }
					}
					closeBlock(writeStartRow + rows.length - 1);
				}
			}

			notifyStatus('Writing to Excel...');

			await Excel.run(async (context) => {
				const sheets = context.workbook.worksheets;
				let sheet = sheets.getItemOrNullObject(targetSheetName);
				await context.sync();
				if (sheet.isNullObject) {
					sheet = sheets.add(targetSheetName);
				}

				const used = sheet.getUsedRangeOrNullObject();
				used.load(['rowCount', 'columnCount']);
				await context.sync();

				if (!used.isNullObject) {
					const existingRows = used.rowCount || 0;
					const existingCols = used.columnCount || 0;
					const rowsToClear = Math.max(0, existingRows - writeStartRow);
					const colsToClear = Math.max(colCount, existingCols, 1);
					if (rowsToClear > 0) {
						const clearRange = sheet.getRangeByIndexes(writeStartRow, 0, rowsToClear, colsToClear);
						clearRange.clear(Excel.ClearApplyTo.contents); 
					}
				}

				if (hasHeaders && sheetHeaders.length > 0) {
					const headerRange = sheet.getRangeByIndexes(0, 0, 1, sheetHeaders.length);
					headerRange.values = [sheetHeaders];
				}

				if (rows.length > 0) {
					const writeRange = sheet.getRangeByIndexes(writeStartRow, 0, rows.length, colCount);
					writeRange.values = rows;
				}

				if (rangesToHighlight.length > 0) {
					queueFormatRanges(sheet, rangesToHighlight, '#ADD8E6'); 
				}
				if (rangesToClear.length > 0) {
					queueFormatRanges(sheet, rangesToClear, null);
				}
                if (cellsToHighlight.length > 0) {
                    queueFormatRanges(sheet, cellsToHighlight, '#ADD8E6'); 
                }

				if (savedStatic && savedStatic.savedMap && savedStatic.savedMap.size > 0) {
					await applyStaticColumnsWithContext(context, sheet, savedStatic, writeStartRow, rowCount, targetSheetName);
				}

				if (conditionalFormat && conditionalFormat.column) {
					const normCol = normalizeKey(conditionalFormat.column);
					const cfColIndex = getHeaderIndexMap(sheetHeaders).get(normCol);
					if (cfColIndex !== undefined && cfColIndex >= 0) {
						applyConditionalFormatting(sheet, cfColIndex, conditionalFormat);
					}
				}

				await context.sync();
			});

			// Update workbook settings with newly discovered columns
			if (newColumnsFromImport.length > 0) {
				try {
					if (typeof window !== 'undefined' && window.Office && Office.context && Office.context.document && Office.context.document.settings) {
						const docSettings = Office.context.document.settings.get('workbookSettings');
						if (docSettings && Array.isArray(docSettings.columns)) {
							const existingNormSet = new Set(docSettings.columns.map(c => normalizeKey(c.name)));
							for (const colName of newColumnsFromImport) {
								if (!existingNormSet.has(normalizeKey(colName))) {
									docSettings.columns.push({ name: colName });
								}
							}
							Office.context.document.settings.set('workbookSettings', docSettings);
							Office.context.document.settings.saveAsync();
							console.log('[DataProcessor] Updated workbook settings with new columns:', newColumnsFromImport);
						}
					}
				} catch (e) {
					console.warn('[DataProcessor] Failed to update workbook settings with new columns', e);
				}
			}

			notifyStatus(`Refresh completed (${rowCount} rows) on ${targetSheetName}.`);
			const result = { success: true, rowsWritten: rowCount, newColumns: newColumnsFromImport };
			if (!suppressCompletion) notifyComplete(result);
			return result;

		} catch (err) {
			console.error('Refresh error', err);
			notifyStatus(`Refresh error: ${String(err)}`);
			const result = { success: false, error: String(err) };
			if (!suppressCompletion) notifyComplete(result);
			return result;
		}
	}

	async function Update(dataToWrite, suppressCompletion = false) {
		notifyStatus('Starting DataProcessor update...');
		
		if (!dataToWrite || !dataToWrite.length) {
			const result = { success: false, reason: 'no-data' };
			if (!suppressCompletion) notifyComplete(result);
			return result;
		}
		
		if (!sheetName || !window.Excel || !window.Excel.run) {
			const result = { success: false, reason: 'missing-requirements' };
			if (!suppressCompletion) notifyComplete(result);
			return result;
		}

		try {
			notifyStatus('Reading current sheet state...');
			
			let usedValues = [];
			let usedFormulas = [];
			let sheetRangeRowIndex = 0;
			let sheetRangeColIndex = 0; 

			await Excel.run(async (context) => {
				const sheets = context.workbook.worksheets;
				const sheet = sheets.getItemOrNullObject(sheetName);
				await context.sync();
				if (sheet.isNullObject) throw new Error('Sheet not found');
				
				const used = sheet.getUsedRangeOrNullObject();
				used.load(['values', 'formulas', 'rowIndex', 'columnIndex', 'rowCount', 'columnCount']);
				await context.sync();
				
				if (!used.isNullObject) {
					usedValues = used.values;
					usedFormulas = used.formulas;
					sheetRangeRowIndex = used.rowIndex;
					sheetRangeColIndex = used.columnIndex || 0;
				}
			});

			if (!usedValues || usedValues.length === 0) {
				notifyStatus('Sheet empty, nothing to update.');
				const result = { success: false, reason: 'empty-sheet' };
				if (!suppressCompletion) notifyComplete(result);
				return result;
			}

			let sheetHeaders = usedValues[0].map(v => (v == null ? '' : String(v).trim()));
			const sheetHeaderIndexMap = getHeaderIndexMap(sheetHeaders);
			
			const dataFirst = dataToWrite[0];
			const isObjectRows = (dataFirst && typeof dataFirst === 'object' && !Array.isArray(dataFirst));
			let dataHeaders = isObjectRows ? Object.keys(dataFirst) : (Array.isArray(dataFirst) ? dataFirst : []);

			const settingsCols = Array.isArray(settingsColumns) ? settingsColumns : [];

			const sheetAliasMap = buildAliasMap(settingsCols, sheetHeaders);
			if (sheetAliasMap.size > 0) {
				sheetHeaders = renameHeaderArray(sheetHeaders, sheetAliasMap);
			}
			const effectiveHeaderIndexMap = getHeaderIndexMap(sheetHeaders);

			const dataAliasMap = buildAliasMap(settingsCols, dataHeaders);
			if (dataAliasMap.size > 0) {
				if (isObjectRows) {
					dataToWrite = renameObjectRows(dataToWrite, dataAliasMap);
				} else if (Array.isArray(dataFirst)) {
					dataToWrite = renameArrayRows(dataToWrite, dataAliasMap);
				}
			}
			
			const fileHeaderKeySet = new Set((isObjectRows ? Object.keys(dataToWrite[0] || {}) : dataHeaders).map(h => normalizeKey(h)));
			
			const matchedSettings = settingsCols.filter(sc => {
				if (!sc.identifier) return false;
				if (fileHeaderKeySet.has(normalizeKey(sc.name))) return true;
				if (sc.alias) {
					const aliases = Array.isArray(sc.alias) ? sc.alias : [sc.alias];
					if (aliases.some(a => fileHeaderKeySet.has(normalizeKey(a)))) return true;
				}
				if (sc.aliases) { 
					if (sc.aliases.some(a => fileHeaderKeySet.has(normalizeKey(a)))) return true;
				}
				return false;
			});
			
			let identifierIndex = -1;
			let identifierSetting = null;
			
			for (const candidate of matchedSettings) {
				const info = computeIdentifierListFromValues(usedValues, sheetHeaders, candidate, usedFormulas);
				if (info.identifierIndex !== -1) {
					identifierIndex = info.identifierIndex;
					identifierSetting = candidate;
					break;
				}
			}

			if (identifierIndex === -1) {
				notifyStatus('Could not match identifier between file and sheet.');
				const result = { success: false, reason: 'no-identifier-match' };
				if (!suppressCompletion) notifyComplete(result);
				return result;
			}

			const normalizedDataKeyMap = new Map();
			if (isObjectRows && dataToWrite.length > 0) {
				Object.keys(dataToWrite[0]).forEach(k => normalizedDataKeyMap.set(normalizeKey(k), k));
			} else if (!isObjectRows && dataToWrite.length > 0) {
				dataToWrite[0].forEach((h, idx) => normalizedDataKeyMap.set(normalizeKey(h), idx));
			}

			let providedHeaders = Array.isArray(matched) ? matched : [];
			const updateCols = []; 
			
			providedHeaders.forEach(ph => {
				let canonicalName = ph;
				const normPH = normalizeKey(ph);
				
				if (dataAliasMap.has(normPH)) {
					canonicalName = dataAliasMap.get(normPH);
				} else if (sheetAliasMap.has(normPH)) {
                    canonicalName = sheetAliasMap.get(normPH);
                }

				const normCanonical = normalizeKey(canonicalName);
				const sheetIdx = effectiveHeaderIndexMap.get(normCanonical);
				const resolvedDataKey = normalizedDataKeyMap.get(normCanonical);

				if (sheetIdx !== undefined && resolvedDataKey !== undefined) {
					if (sheetIdx === identifierIndex) {
						return;
					}
					updateCols.push({ 
						index: sheetIdx, 
						dataLookup: resolvedDataKey,
						name: canonicalName
					});
				}
			});

			if (updateCols.length === 0) {
				const result = { success: false, reason: 'no-update-columns' };
				if (!suppressCompletion) notifyComplete(result);
				return result;
			}

			const updatesByColIndex = new Map(); 

			for (const col of updateCols) {
				const colData = usedFormulas.map(row => row[col.index]); 
				updatesByColIndex.set(col.index, colData);
			}

			const idToRowMap = new Map();
			for (let r = 1; r < usedValues.length; r++) {
				let val = usedValues[r][identifierIndex];
				if (usedFormulas[r][identifierIndex] && String(usedFormulas[r][identifierIndex]).startsWith('=')) {
					const link = extractHyperlink(usedFormulas[r][identifierIndex]);
					if (link) val = link;
				}
				const id = String(val || '').trim();
				if (id) idToRowMap.set(id, r);
			}

			let updatedCount = 0;
			let skippedCount = 0;

			const idKeyNorm = normalizeKey(identifierSetting.name);
			let idLookup = normalizedDataKeyMap.get(idKeyNorm);
			
			if (idLookup === undefined && identifierSetting.alias) {
				const aliases = Array.isArray(identifierSetting.alias) ? identifierSetting.alias : [identifierSetting.alias];
				for (const a of aliases) {
					const k = normalizedDataKeyMap.get(normalizeKey(a));
					if (k !== undefined) {
						idLookup = k;
						break;
					}
				}
			}

			for (const rowIn of dataToWrite) {
				let incomingId = '';
				if (idLookup !== undefined) {
					incomingId = String(rowIn[idLookup] || '').trim();
				}
				
				if (!incomingId || !idToRowMap.has(incomingId)) {
					skippedCount++;
					continue;
				}

				const rowIndex = idToRowMap.get(incomingId);
				updatedCount++;

				for (const col of updateCols) {
					const colArr = updatesByColIndex.get(col.index);
					const newVal = rowIn[col.dataLookup] !== undefined ? rowIn[col.dataLookup] : '';
					colArr[rowIndex] = newVal;
				}
			}

			if (updatedCount === 0) {
				notifyStatus(`Update finished but 0 rows matched. Checked ${dataToWrite.length} rows, skipped ${skippedCount}.`);
				const result = { success: true, rowsUpdated: 0, rowsSkipped: skippedCount, warning: 'No rows matched' };
				if (!suppressCompletion) notifyComplete(result);
				return result;
			}

			notifyStatus(`Applying ${updatedCount} row updates across ${updateCols.length} columns...`);
			
			await Excel.run(async (context) => {
				const sheets = context.workbook.worksheets;
				const sheet = sheets.getItemOrNullObject(sheetName);
				await context.sync();
				
				for (const [colIdx, newColData] of updatesByColIndex.entries()) {
					const colMatrix = newColData.map(v => [v]);
					const absoluteColIndex = sheetRangeColIndex + colIdx;
					const rng = sheet.getRangeByIndexes(sheetRangeRowIndex, absoluteColIndex, colMatrix.length, 1);
					rng.formulas = colMatrix; 
				}

				if (conditionalFormat && conditionalFormat.column) {
					const normCol = normalizeKey(conditionalFormat.column);
					const cfColIndex = effectiveHeaderIndexMap.get(normCol);
					
					if (cfColIndex !== undefined && cfColIndex >= 0) {
						const absColIndex = sheetRangeColIndex + cfColIndex;
						applyConditionalFormatting(sheet, absColIndex, conditionalFormat);
					}
				}

				await context.sync();
			});

			notifyStatus(`Update complete.`);
			const result = { success: true, rowsUpdated: updatedCount, rowsSkipped: skippedCount };
			if (!suppressCompletion) notifyComplete(result);
			return result;

		} catch (err) {
			console.error('Update error', err);
			notifyStatus(`Update error: ${String(err)}`);
			const result = { success: false, error: String(err) };
			if (!suppressCompletion) notifyComplete(result);
			return result;
		}
	}

	async function Hybrid(dataToWrite) {
		if (!refreshSheetName) {
			notifyStatus('Hybrid action failed: No refresh sheet specified.');
			notifyComplete({ success: false, reason: 'missing-refresh-sheet' });
			return;
		}

		notifyStatus(`Hybrid Step 1: Refreshing "${refreshSheetName}"...`);
		const refreshResult = await Refresh(dataToWrite, refreshSheetName, true); 

		if (!refreshResult || !refreshResult.success) {
			notifyStatus(`Hybrid failed during Refresh phase: ${refreshResult?.error || refreshResult?.reason}`);
			notifyComplete(refreshResult);
			return;
		}

		notifyStatus(`Hybrid Step 2: Updating "${sheetName}"...`);
		const updateResult = await Update(dataToWrite, true); 

		const finalResult = {
			success: updateResult.success,
			refreshResult: refreshResult,
			updateResult: updateResult,
			rowsWritten: refreshResult.rowsWritten,
			rowsUpdated: updateResult.rowsUpdated,
			rowsSkipped: updateResult.rowsSkipped
		};

		if (updateResult.success) {
			notifyStatus('Hybrid action completed successfully.');
		} else {
			notifyStatus(`Hybrid failed during Update phase: ${updateResult.error || updateResult.reason}`);
		}
		
		notifyComplete(finalResult);
	}

	useEffect(() => {
		if (action === 'Refresh') {
			Refresh(data, sheetName, false); 
		} else if (action === 'Update') {
			Update(data, false); 
		} else if (action === 'Hybrid') {
			Hybrid(data);
		} else {
			if (action && action !== 'None') notifyStatus(`Unknown action: ${action}`);
		}
	}, [data, sheetName, refreshSheetName, action]);

	return null;
}