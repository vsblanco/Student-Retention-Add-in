// [2025-11-19] v2.1 - Exclude Identifier from Update
// Changes:
// - In Update(), filtered out the identifier column from the list of columns to update.
// - This prevents overwriting the ID cell with the same value, which is redundant.

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

export default function DataProcessor({ data, sheetName, refreshSheetName, settingsColumns, matched, onComplete, onStatus, action }) {
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

	// Helper to queue formatting commands for a list of addresses.
	function queueFormatRanges(sheet, addresses, color) {
		if (!addresses || addresses.length === 0) return;
		for (const addr of addresses) {
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
		}
	}

	// Refresh: write `data` to the target worksheet.
	// Added suppressCompletion so it can be reused by Hybrid without ending the process early.
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
			let sheetExisted = false;

			await Excel.run(async (context) => {
				const sheets = context.workbook.worksheets;
				const sheet = sheets.getItemOrNullObject(targetSheetName);
				await context.sync();
				if (!sheet.isNullObject) {
					sheetExisted = true;
					const used = sheet.getUsedRangeOrNullObject();
					used.load(['values', 'formulas']);
					await context.sync();
					if (!used.isNullObject) {
						usedValues = used.values || [];
						usedFormulas = used.formulas || [];
					}
				}
			});

			notifyStatus('Processing data structure...');

			// derived headers
			let sheetHeaders = (sheetExisted && usedValues.length > 0)
				? usedValues[0].map(v => (v == null ? '' : String(v).trim()))
				: [];
			const hasHeaders = sheetHeaders.some(h => h !== '');
			const sheetHeaderIndexMap = getHeaderIndexMap(sheetHeaders);

			// incoming data headers
			const dataFirst = dataToWrite[0];
			let dataHeaders = (dataFirst && typeof dataFirst === 'object' && !Array.isArray(dataFirst))
				? Object.keys(dataFirst).map(h => (h == null ? '' : String(h).trim()))
				: (Array.isArray(dataFirst) ? dataFirst.map(h => (h == null ? '' : String(h).trim())) : []);
			const dataHeaderIndexMap = getHeaderIndexMap(dataHeaders);

			const settingsCols = Array.isArray(settingsColumns) ? settingsColumns : [];
			const identifierSetting = gatherIdentifierColumn(settingsCols);

			// Capture PRE-EXISTING identifiers for diffing
			let beforeIdentifiers = new Set();
			if (identifierSetting && usedValues.length > 0) {
				const beforeInfo = computeIdentifierListFromValues(usedValues, sheetHeaders, identifierSetting, usedFormulas);
				beforeIdentifiers = beforeInfo.identifiers;
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

			// Apply aliases to headers
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

			// Normalize Data Rows
			const combinedDataAliasMap = new Map([...matchedByName, ...dataAliasMap]);
			if (combinedDataAliasMap.size > 0) {
				const firstRow = dataToWrite[0];
				if (firstRow && typeof firstRow === 'object' && !Array.isArray(firstRow)) {
					dataToWrite = renameObjectRows(dataToWrite, combinedDataAliasMap);
				} else if (Array.isArray(firstRow)) {
					dataToWrite = renameArrayRows(dataToWrite, combinedDataAliasMap);
				}
			}

			// Static Columns Calculation
			let dataFieldSet = new Set();
			const firstAfter = dataToWrite[0];
			if (firstAfter && typeof firstAfter === 'object' && !Array.isArray(firstAfter)) {
				dataToWrite.forEach(row => Object.keys(row || {}).forEach(k => dataFieldSet.add(normalizeKey(k))));
			} else if (Array.isArray(firstAfter)) {
				firstAfter.forEach(v => dataFieldSet.add(normalizeKey(v)));
			}

			const staticCols = [];
			sheetHeaders.forEach(h => {
				const key = normalizeKey(h);
				if (key && !dataFieldSet.has(key)) staticCols.push(h);
			});

			let savedStatic = null;
			if (staticCols.length > 0 && identifierSetting) {
				savedStatic = computeSavedStaticFromValues(usedValues, sheetHeaders, staticCols, identifierSetting, usedFormulas);
			}

			// Prepare Matrix for Writing
			notifyStatus('Constructing write matrix...');
			const isObjectRows = (firstAfter && typeof firstAfter === 'object' && !Array.isArray(firstAfter));
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
				// No headers, just dump data
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

			// PRE-CALCULATE HIGHLIGHTING RANGES (In-Memory)
			const rangesToHighlight = [];
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

						const type = !beforeIdentifiers.has(id) ? 'highlight' : 'clear';
						const absRowIndex = writeStartRow + r;

						if (type !== currentBlockType) {
							closeBlock(absRowIndex - 1);
							currentBlockType = type;
							currentBlockStart = absRowIndex;
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

				// Clear existing data (below header if exists)
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

				// Write Headers
				if (hasHeaders && sheetHeaders.length > 0) {
					const headerRange = sheet.getRangeByIndexes(0, 0, 1, sheetHeaders.length);
					headerRange.values = [sheetHeaders];
				}

				// Write Data
				if (rows.length > 0) {
					const writeRange = sheet.getRangeByIndexes(writeStartRow, 0, rows.length, colCount);
					writeRange.values = rows;
				}

				// Restore Static Columns
				if (savedStatic && savedStatic.savedMap && savedStatic.savedMap.size > 0) {
					await applyStaticColumnsWithContext(context, sheet, savedStatic, writeStartRow, rowCount, targetSheetName);
				}

				// Apply Highlighting
				if (rangesToHighlight.length > 0) {
					queueFormatRanges(sheet, rangesToHighlight, 'lightblue');
				}
				if (rangesToClear.length > 0) {
					queueFormatRanges(sheet, rangesToClear, null);
				}

				await context.sync();
			});

			notifyStatus(`Refresh completed (${rowCount} rows) on ${targetSheetName}.`);
			const result = { success: true, rowsWritten: rowCount };
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

	// Update: update matched columns in existing rows
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
			
			// 1. PREPARE & READ
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

			// Derive Headers and Normalize Data
			let sheetHeaders = usedValues[0].map(v => (v == null ? '' : String(v).trim()));
			const sheetHeaderIndexMap = getHeaderIndexMap(sheetHeaders);
			
			// Identify headers in incoming data
			const dataFirst = dataToWrite[0];
			const isObjectRows = (dataFirst && typeof dataFirst === 'object' && !Array.isArray(dataFirst));
			let dataHeaders = isObjectRows ? Object.keys(dataFirst) : (Array.isArray(dataFirst) ? dataFirst : []);

			const settingsCols = Array.isArray(settingsColumns) ? settingsColumns : [];

			// 2. NORMALIZE HEADERS (Decoupled Sheet vs Data)
			
			// A. Normalize SHEET Headers (Rename Alias -> Canonical)
			const sheetAliasMap = buildAliasMap(settingsCols, sheetHeaders);
			if (sheetAliasMap.size > 0) {
				sheetHeaders = renameHeaderArray(sheetHeaders, sheetAliasMap);
			}
			const effectiveHeaderIndexMap = getHeaderIndexMap(sheetHeaders);

			// B. Normalize DATA Headers (Rename Alias -> Canonical)
			// We check ALL settings against Data, regardless of whether they matched on the Sheet.
			const dataAliasMap = buildAliasMap(settingsCols, dataHeaders);
			if (dataAliasMap.size > 0) {
				if (isObjectRows) {
					dataToWrite = renameObjectRows(dataToWrite, dataAliasMap);
				} else if (Array.isArray(dataFirst)) {
					dataToWrite = renameArrayRows(dataToWrite, dataAliasMap);
				}
			}
			
			// 3. Find Identifier in File
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
				// Check if this candidate exists on the SHEET
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

			// 4. Map Update Columns
			const normalizedDataKeyMap = new Map();
			if (isObjectRows && dataToWrite.length > 0) {
				Object.keys(dataToWrite[0]).forEach(k => normalizedDataKeyMap.set(normalizeKey(k), k));
			} else if (!isObjectRows && dataToWrite.length > 0) {
				dataToWrite[0].forEach((h, idx) => normalizedDataKeyMap.set(normalizeKey(h), idx));
			}

			let providedHeaders = Array.isArray(matched) ? matched : [];
			const updateCols = []; // { index, dataKey/dataIndex, name }
			
			providedHeaders.forEach(ph => {
				// Resolve 'ph' (File Header) to 'Canonical Name' via Data Map
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
					// --- NEW CHANGE v2.1: Skip Identifier Column ---
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

			// 5. BUILD UPDATES (IN MEMORY)
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

			// Resolve identifier key for incoming data
			const idKeyNorm = normalizeKey(identifierSetting.name);
			let idLookup = normalizedDataKeyMap.get(idKeyNorm);
			
			// Fallback: Check alias keys if main key missing in data
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

			// Iterate incoming data
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

			// 6. WRITE BACK
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

	// Hybrid: Refresh separate sheet -> Update current sheet
	async function Hybrid(dataToWrite) {
		if (!refreshSheetName) {
			notifyStatus('Hybrid action failed: No refresh sheet specified.');
			notifyComplete({ success: false, reason: 'missing-refresh-sheet' });
			return;
		}

		notifyStatus(`Hybrid Step 1: Refreshing "${refreshSheetName}"...`);
		const refreshResult = await Refresh(dataToWrite, refreshSheetName, true); // Suppress completion

		if (!refreshResult || !refreshResult.success) {
			notifyStatus(`Hybrid failed during Refresh phase: ${refreshResult?.error || refreshResult?.reason}`);
			notifyComplete(refreshResult);
			return;
		}

		notifyStatus(`Hybrid Step 2: Updating "${sheetName}"...`);
		const updateResult = await Update(dataToWrite, true); // Suppress completion

		// Merge results for final notification
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
			Refresh(data, sheetName, false); // Normal refresh, target = sheetName
		} else if (action === 'Update') {
			Update(data, false); // Normal update
		} else if (action === 'Hybrid') {
			Hybrid(data);
		} else {
			if (action && action !== 'None') notifyStatus(`Unknown action: ${action}`);
		}
	}, [data, sheetName, refreshSheetName, action]);

	return null;
}