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
	getHeaderIndexMap, // added for fast header lookups
} from './dataProcessorUtility';

export default function DataProcessor({ data, sheetName, settingsColumns, headers, onComplete, onStatus }) {
	// helper to notify caller (guarded)
	const notifyComplete = (payload) => {
		if (typeof onComplete === 'function') {
			try { onComplete(payload); } catch (e) { /* swallow callback errors */ }
		}
	};

	// new: helper to send incremental status updates
	const notifyStatus = (msg) => {
		if (typeof onStatus === 'function') {
			try { onStatus(String(msg)); } catch (e) { /* swallow callback errors */ }
		}
	};

	// Refresh: write `data` to the worksheet named `sheetName` using the Excel JS API.
	async function Refresh(dataToWrite) {
		notifyStatus('Starting DataProcessor refresh...');
		if (!dataToWrite || !dataToWrite.length) {
			console.warn('Refresh: no data to write');
			notifyStatus('No data to write.');
			notifyComplete({ success: false, reason: 'no-data' });
			return;
		}
		if (!sheetName) {
			console.warn('Refresh: missing sheetName');
			notifyStatus('Missing sheet name.');
			notifyComplete({ success: false, reason: 'missing-sheetName' });
			return;
		}
		if (!window.Excel || !window.Excel.run) {
			console.error('Excel JS API is not available in this context.');
			notifyStatus('Excel JS API not available.');
			notifyComplete({ success: false, reason: 'no-excel-api' });
			return;
		}

		let rowCount = 0; // visible to final log
		try {
			notifyStatus('Reading existing worksheet (if any)...');
			// INITIAL READ: single Excel.run to obtain existing used values + formulas (if sheet exists)
			let usedValues = [];
			let usedFormulas = [];
			let sheetExisted = false;
			await Excel.run(async (context) => {
				const sheets = context.workbook.worksheets;
				const sheet = sheets.getItemOrNullObject(sheetName);
				await context.sync();
				if (sheet.isNullObject) {
					sheetExisted = false;
					usedValues = [];
				} else {
					sheetExisted = true;
					const used = sheet.getUsedRangeOrNullObject();
					used.load(['values', 'formulas']);
					await context.sync();
					if (!used.isNullObject) {
						usedValues = Array.isArray(used.values) ? used.values : [];
						usedFormulas = Array.isArray(used.formulas) ? used.formulas : [];
					} else {
						usedValues = [];
					}
				}
			});

			notifyStatus('Derived existing sheet headers and formulas.');

			// derive sheetHeaders (trim once) and index map for O(1) membership/index lookups
			let sheetHeaders = Array.isArray(usedValues) && usedValues[0]
				? usedValues[0].map((v) => (v == null ? '' : String(v).trim()))
				: [];
			const hasHeaders = Array.isArray(sheetHeaders) && sheetHeaders.some((h) => h !== '');
			const sheetHeaderIndexMap = getHeaderIndexMap(sheetHeaders);

			// derive data headers from incoming data and build index map
			const dataFirst = dataToWrite[0];
			let dataHeaders = (dataFirst && typeof dataFirst === 'object' && !Array.isArray(dataFirst))
				? Object.keys(dataFirst).map((h) => (h == null ? '' : String(h).trim()))
				: (Array.isArray(dataFirst) ? dataFirst.map((h) => (h == null ? '' : String(h).trim())) : []);
			const dataHeaderIndexMap = getHeaderIndexMap(dataHeaders);

			// normalize settings and pick identifier early
			const settingsCols = Array.isArray(settingsColumns) ? settingsColumns : [];
			const identifierSetting = gatherIdentifierColumn(settingsCols);

			// CAPTURE BEFORE IDENTIFIERS
			let beforeIdentifiers = null;
			if (identifierSetting && Array.isArray(usedValues) && usedValues.length > 0) {
				const beforeInfo = computeIdentifierListFromValues(usedValues, sheetHeaders, identifierSetting, usedFormulas);
				beforeIdentifiers = beforeInfo.identifiers;
			}

			notifyStatus('Normalizing and mapping incoming data columns...');
			// 1) Match canonical names by checking the header-index maps (O(1) checks)
			const matchedByName = new Map();
			for (let i = 0; i < settingsCols.length; i++) {
				const sc = settingsCols[i];
				if (!sc || !sc.name) continue;
				const nameKey = normalizeKey(sc.name);
				if (sheetHeaderIndexMap.has(nameKey) && dataHeaderIndexMap.has(nameKey)) {
					matchedByName.set(nameKey, String(sc.name).trim());
				}
			}

			// 2) For remaining settings, run alias checks
			const remainingSettings = settingsCols.filter((sc) => sc && sc.name && !matchedByName.has(normalizeKey(sc.name)));
			const sheetAliasMap = buildAliasMap(remainingSettings, sheetHeaders);
			const dataAliasMap = buildAliasMap(remainingSettings, dataHeaders);

			// 3) Canonicalize exact-name matches on worksheet headers (in-memory)
			if (matchedByName.size > 0) {
				sheetHeaders = sheetHeaders.map((h) => {
					const key = normalizeKey(h);
					return matchedByName.has(key) ? matchedByName.get(key) : h;
				});
				// rebuild index map after possible canonicalization
				// (only needed for subsequent logic)
				// note: getHeaderIndexMap is cheap and cached via normalizeKey cache
				Object.assign(sheetHeaderIndexMap, getHeaderIndexMap(sheetHeaders));
			}

			// 4) Apply alias renames to worksheet headers in-memory
			if (sheetAliasMap && sheetAliasMap.size > 0) {
				sheetHeaders = renameHeaderArray(sheetHeaders, sheetAliasMap);
			}

			// 5) Normalize incoming data column names to canonical names
			const combinedDataAliasMap = new Map();
			for (const [k, v] of matchedByName.entries()) combinedDataAliasMap.set(k, v);
			if (dataAliasMap && dataAliasMap.size > 0) {
				for (const [k, v] of dataAliasMap.entries()) combinedDataAliasMap.set(k, v);
			}

			if (combinedDataAliasMap && combinedDataAliasMap.size > 0) {
				const firstRow = dataToWrite[0];
				if (firstRow && typeof firstRow === 'object' && !Array.isArray(firstRow)) {
					dataToWrite = renameObjectRows(dataToWrite, combinedDataAliasMap);
				} else if (Array.isArray(firstRow)) {
					dataToWrite = renameArrayRows(dataToWrite, combinedDataAliasMap);
				}
			}

			// DETERMINE static columns: present on worksheet but not in data
			let dataFieldSet = new Set();
			const firstAfter = dataToWrite && dataToWrite[0];
			const isObjectRowsAfter = firstAfter && typeof firstAfter === 'object' && !Array.isArray(firstAfter);
			if (isObjectRowsAfter) {
				for (let i = 0; i < dataToWrite.length; i++) {
					const rowObj = dataToWrite[i];
					if (!rowObj || typeof rowObj !== 'object' || Array.isArray(rowObj)) continue;
					const keys = Object.keys(rowObj);
					for (let k = 0; k < keys.length; k++) {
						const key = normalizeKey(keys[k]);
						if (key) dataFieldSet.add(key);
					}
				}
			} else if (Array.isArray(firstAfter)) {
				for (let i = 0; i < firstAfter.length; i++) dataFieldSet.add(normalizeKey(firstAfter[i]));
			}

			const sheetHeadersNormalized = sheetHeaders.map((v) => (v == null ? '' : String(v).trim()));
			const staticCols = [];
			for (let i = 0; i < sheetHeadersNormalized.length; i++) {
				const h = sheetHeadersNormalized[i];
				const key = normalizeKey(h);
				if (key && !dataFieldSet.has(key)) staticCols.push(h);
			}

			// capture static values (if needed)
			let savedStatic = null;
			if (staticCols.length > 0) {
				if (!identifierSetting || !identifierSetting.name) {
					console.warn('No identifier configured; skipping static column retention.');
				} else {
					savedStatic = computeSavedStaticFromValues(usedValues, sheetHeaders, staticCols, identifierSetting, usedFormulas);
				}
			}

			notifyStatus('Preparing rows to write...');
			// Build rows to write and decide writeStartRow (minimize intermediate allocations)
			const first = dataToWrite[0];
			const isObjectRows = first && typeof first === 'object' && !Array.isArray(first);

			let rows;
			let writeStartRow = 0;

			if (hasHeaders) {
				const desiredColumns = sheetHeaders
					.map((h, idx) => (h ? { name: h, index: idx } : null))
					.filter(Boolean);

				if (isObjectRows) {
					rows = new Array(dataToWrite.length);
					for (let i = 0; i < dataToWrite.length; i++) {
						const obj = dataToWrite[i];
						const out = new Array(desiredColumns.length);
						for (let j = 0; j < desiredColumns.length; j++) {
							const c = desiredColumns[j];
							out[j] = (obj && Object.prototype.hasOwnProperty.call(obj, c.name)) ? obj[c.name] : '';
						}
						rows[i] = out;
					}
				} else {
					rows = new Array(dataToWrite.length);
					for (let i = 0; i < dataToWrite.length; i++) {
						const arr = dataToWrite[i];
						const out = new Array(desiredColumns.length);
						for (let j = 0; j < desiredColumns.length; j++) {
							const c = desiredColumns[j];
							out[j] = (Array.isArray(arr) && arr.length > c.index) ? arr[c.index] : '';
						}
						rows[i] = out;
					}
				}
				writeStartRow = 1;
			} else {
				if (isObjectRows) {
					const headers = Object.keys(first);
					const headerRow = headers.slice();
					rows = new Array(dataToWrite.length + 1);
					rows[0] = headerRow;
					for (let i = 0; i < dataToWrite.length; i++) {
						const obj = dataToWrite[i];
						const out = new Array(headers.length);
						for (let j = 0; j < headers.length; j++) out[j] = obj[headers[j]] ?? '';
						rows[i + 1] = out;
					}
					writeStartRow = 0;
				} else {
					rows = dataToWrite.slice();
					writeStartRow = 0;
				}
			}

			rowCount = rows.length;
			let colCount = rows[0] ? rows[0].length : 0;
			if (rowCount === 0 || colCount === 0) {
				console.warn('Refresh: nothing to write after normalization');
				notifyComplete({ success: false, reason: 'nothing-to-write' });
				return;
			}

			// normalize row widths once
			for (let i = 0; i < rows.length; i++) {
				const r = rows[i];
				if (!Array.isArray(r)) rows[i] = new Array(colCount).fill('');
				else if (r.length < colCount) {
					const fill = new Array(colCount - r.length).fill('');
					rows[i] = r.concat(fill);
				} else if (r.length > colCount) {
					rows[i] = r.slice(0, colCount);
				}
			}

			notifyStatus('Writing data to worksheet...');
			// FINAL WRITE: single Excel.run to create sheet (if needed), clear below header, write header + data, restore static columns
			await Excel.run(async (context) => {
				const sheets = context.workbook.worksheets;
				let sheet = sheets.getItemOrNullObject(sheetName);
				await context.sync();
				if (sheet.isNullObject) {
					sheet = sheets.add(sheetName);
				}

				// clear existing rows below header (minimize writes)
				const used = sheet.getUsedRangeOrNullObject();
				used.load(['rowCount', 'columnCount']);
				await context.sync();

				if (!used.isNullObject) {
					const existingRowCount = used.rowCount || 0;
					const existingColCount = used.columnCount || 0;
					const rowsToClear = Math.max(0, existingRowCount - writeStartRow);
					const colsToClear = Math.max(colCount, existingColCount, 1);
					if (rowsToClear > 0 && colsToClear > 0) {
						const blankRow = new Array(colsToClear).fill('');
						const blankMatrix = new Array(rowsToClear);
						for (let i = 0; i < rowsToClear; i++) blankMatrix[i] = blankRow.slice();
						const clearRange = sheet.getRangeByIndexes(writeStartRow, 0, rowsToClear, colsToClear);
						clearRange.values = blankMatrix;
					}
				}

				// write header row if present
				if (hasHeaders && Array.isArray(sheetHeaders) && sheetHeaders.length > 0) {
					const headerRange = sheet.getRangeByIndexes(0, 0, 1, sheetHeaders.length);
					headerRange.values = [sheetHeaders];
				}

				// bulk-write rows
				const writeRange = sheet.getRangeByIndexes(writeStartRow, 0, rowCount, colCount);
				writeRange.values = rows;

				// restore static columns inside same context
				if (savedStatic && savedStatic.savedMap && savedStatic.savedMap.size > 0) {
					await applyStaticColumnsWithContext(context, sheet, savedStatic, writeStartRow, rowCount, sheetName);
				}

				await context.sync();
			});

			notifyStatus('Data written; applying static columns and finalizing...');
			// AFTER WRITE: read used-range again and perform diff + clear/highlight in a single Excel.run (reduces round trips)
			if (identifierSetting && beforeIdentifiers) {
				notifyStatus('Detecting/highlighting new identifiers...');
				try {
					await Excel.run(async (context) => {
						const sheets = context.workbook.worksheets;
						const sheet = sheets.getItemOrNullObject(sheetName);
						await context.sync();
						if (sheet.isNullObject) return;

						const used = sheet.getUsedRangeOrNullObject();
						used.load(['values', 'formulas', 'rowIndex', 'columnIndex', 'columnCount', 'isNullObject']);
						await context.sync();
						if (used.isNullObject) return;

						const afterValues = used.values || [];
						const afterFormulas = Array.isArray(used.formulas) ? used.formulas : [];
						const afterInfo = computeIdentifierListFromValues(afterValues, sheetHeaders, identifierSetting, afterFormulas);
						const afterIdentifiers = afterInfo.identifiers;

						// compute diffs
						const newIds = [];
						const existingIds = [];
						for (const id of afterIdentifiers) {
							if (!beforeIdentifiers.has(id)) newIds.push(id);
							else existingIds.push(id);
						}
						if (newIds.length === 0 && existingIds.length === 0) return;

						const colCountUsed = used.columnCount || (afterValues[0] ? afterValues[0].length : 0);
						const startCol = (typeof used.columnIndex === 'number') ? used.columnIndex : 0;
						const rightmostCol = startCol + Math.max(colCountUsed, 1) - 1;
						const leftStartCol = Math.max(0, rightmostCol - Math.max(colCountUsed, 1) + 1);
						const actualColCount = (rightmostCol >= leftStartCol) ? (rightmostCol - leftStartCol + 1) : Math.max(colCountUsed, 1);

						// locate absolute row indices for each id and queue format changes
						for (let r = 1; r < afterValues.length; r++) {
							const row = Array.isArray(afterValues[r]) ? afterValues[r] : [];
							let raw = row[afterInfo.identifierIndex];
							if (afterFormulas && afterFormulas[r] && typeof afterFormulas[r][afterInfo.identifierIndex] === 'string' && afterFormulas[r][afterInfo.identifierIndex].trim().startsWith('=')) {
								const link = extractHyperlink(afterFormulas[r][afterInfo.identifierIndex]);
								if (link) raw = link;
							}
							const id = raw == null ? '' : String(raw).trim();
							if (newIds.includes(id)) {
								const absRow = (typeof used.rowIndex === 'number' ? used.rowIndex : 0) + r;
								try {
									const rng = sheet.getRangeByIndexes(absRow, leftStartCol, 1, actualColCount);
									rng.format.fill.color = 'lightblue';
								} catch (inner) {
									console.error('queue highlight range failed for', id, inner);
								}
							} else if (existingIds.includes(id)) {
								const absRow = (typeof used.rowIndex === 'number' ? used.rowIndex : 0) + r;
								try {
									const rng = sheet.getRangeByIndexes(absRow, leftStartCol, 1, actualColCount);
									rng.format.fill.clear();
								} catch (inner) {
									console.error('queue clear range failed for', id, inner);
								}
							}
						}

						await context.sync();
					});
				} catch (err) {
					console.error('Error detecting/highlighting new identifiers after import', err);
				}
				notifyStatus('Identifier highlighting complete.');
			}

			// final successful write
			console.log(`Refresh concluded for sheet="${sheetName}", rowsWritten=${rowCount}`);
			notifyStatus(`Refresh completed (${rowCount} rows).`);
			notifyComplete({ success: true, rowsWritten: rowCount });
		} catch (err) {
			console.error('Refresh error', err);
			notifyStatus(`Refresh error: ${String(err)}`);
			notifyComplete({ success: false, error: String(err) });
		}
	}

	useEffect(() => {
		// call Refresh when data or sheetName changes
		if (data && sheetName) {
			Refresh(data);
		}
	}, [data, sheetName]);

	return null;
}
