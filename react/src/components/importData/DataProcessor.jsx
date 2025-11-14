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

export default function DataProcessor({ data, sheetName, settingsColumns, matched, onComplete, onStatus, action }) {
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

			// Log all settings entries that are marked as identifiers (flatten aliases)
			try {
				const identifierEntries = (settingsCols || []).filter((sc) => sc && sc.identifier);
				const idInfo = identifierEntries.map((sc) => {
					// collect alias sources into an array (may be string or array)
					let rawAliases = [];
					if (Array.isArray(sc.aliases)) rawAliases = sc.aliases.slice();
					else if (Array.isArray(sc.alias)) rawAliases = sc.alias.slice();
					else if (typeof sc.alias === 'string') rawAliases = [sc.alias];

					// fallback alternate properties
					if ((!rawAliases || rawAliases.length === 0) && Array.isArray(sc.alternates)) rawAliases = sc.alternates.slice();
					if ((!rawAliases || rawAliases.length === 0) && sc.alt) rawAliases = Array.isArray(sc.alt) ? sc.alt.slice() : [sc.alt];

					// flatten any nested arrays and remove null/undefined
					rawAliases = rawAliases.reduce((acc, v) => {
						if (Array.isArray(v)) return acc.concat(v.filter(Boolean));
						if (v == null) return acc;
						return acc.concat([v]);
					}, []);

					const aliasesNormalized = rawAliases.map((a) => (a ? normalizeKey(String(a)) : '')).filter(Boolean);
					return {
						name: sc.name || '',
						normalized: normalizeKey(sc.name || ''),
						aliases: rawAliases,
						aliasesNormalized,
						raw: sc,
					};
				});
				console.log('Settings identifier entries (update):', idInfo);
				notifyStatus(`Settings identifiers (update): ${idInfo.map((i) => i.name || '(unnamed)').join(', ') || '(none)'}`);
			} catch (logErr) {
				console.warn('Failed logging identifier entries from settings (update)', logErr);
			}
			// Log workbook settings and identifier info for debugging
			try {
				const settingsNames = settingsCols.map((sc) => (sc && sc.name ? String(sc.name).trim() : ''));
				console.log('Workbook settings columns:', settingsNames);
				notifyStatus(`Workbook settings columns: ${settingsNames.join(', ') || '(none)'}`);
				console.log('gatherIdentifierColumn ->', identifierSetting);
				notifyStatus(`Identifier setting from workbook: ${identifierSetting && identifierSetting.name ? identifierSetting.name : 'none'}`);
			} catch (logErr) {
				console.warn('Failed logging workbook settings/identifier', logErr);
			}
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

	async function Update(dataToWrite) {
		notifyStatus('Starting DataProcessor update (no-remove mode)...');
		if (!dataToWrite || !dataToWrite.length) {
			console.warn('Update: no data to update');
			notifyStatus('No data to update.');
			notifyComplete({ success: false, reason: 'no-data' });
			return;
		}
		if (!sheetName) {
			console.warn('Update: missing sheetName');
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

		try {
			notifyStatus('Reading existing worksheet for update...');
			let usedValues = [];
			let usedFormulas = [];
			let sheetRangeRowIndex = 0;
			let sheetRowCount = 0;
			let sheetColCount = 0;
			await Excel.run(async (context) => {
				const sheets = context.workbook.worksheets;
				const sheet = sheets.getItemOrNullObject(sheetName);
				await context.sync();
				if (sheet.isNullObject) {
					notifyStatus('Sheet not found for update.');
					notifyComplete({ success: false, reason: 'sheet-not-found' });
					return;
				}
				const used = sheet.getUsedRangeOrNullObject();
				used.load(['values', 'formulas', 'rowIndex', 'rowCount', 'columnCount', 'isNullObject']);
				await context.sync();
				if (used.isNullObject) {
					notifyStatus('Sheet is empty; performing append as update.');
				} else {
					usedValues = Array.isArray(used.values) ? used.values : [];
					usedFormulas = Array.isArray(used.formulas) ? used.formulas : [];
					sheetRangeRowIndex = typeof used.rowIndex === 'number' ? used.rowIndex : 0;
					sheetRowCount = typeof used.rowCount === 'number' ? used.rowCount : (usedValues.length || 0);
					sheetColCount = typeof used.columnCount === 'number' ? used.columnCount : (usedValues[0] ? usedValues[0].length : 0);
				}
			});

			// derive sheetHeaders if present
			const sheetHeaders = Array.isArray(usedValues) && usedValues[0]
				? usedValues[0].map((v) => (v == null ? '' : String(v).trim()))
				: [];

			// prepare header index maps and normalize incoming data (reuse same logic as Refresh)
			const dataFirst = dataToWrite[0];
			let dataHeaders = (dataFirst && typeof dataFirst === 'object' && !Array.isArray(dataFirst))
				? Object.keys(dataFirst).map((h) => (h == null ? '' : String(h).trim()))
				: (Array.isArray(dataFirst) ? dataFirst.map((h) => (h == null ? '' : String(h).trim())) : []);
			const sheetHeaderIndexMap = getHeaderIndexMap(sheetHeaders);
			const dataHeaderIndexMap = getHeaderIndexMap(dataHeaders);
			const settingsCols = Array.isArray(settingsColumns) ? settingsColumns : [];

			// Log all settings entries that are marked as identifiers (update flow)
			try {
				const identifierEntries = (settingsCols || []).filter((sc) => sc && sc.identifier);
				const idInfo = identifierEntries.map((sc) => {
					// collect alias sources into an array (may be string or array)
					let rawAliases = [];
					if (Array.isArray(sc.aliases)) rawAliases = sc.aliases.slice();
					else if (Array.isArray(sc.alias)) rawAliases = sc.alias.slice();
					else if (typeof sc.alias === 'string') rawAliases = [sc.alias];

					// fallback alternate properties
					if ((!rawAliases || rawAliases.length === 0) && Array.isArray(sc.alternates)) rawAliases = sc.alternates.slice();
					if ((!rawAliases || rawAliases.length === 0) && sc.alt) rawAliases = Array.isArray(sc.alt) ? sc.alt.slice() : [sc.alt];

					// flatten any nested arrays and remove null/undefined
					rawAliases = rawAliases.reduce((acc, v) => {
						if (Array.isArray(v)) return acc.concat(v.filter(Boolean));
						if (v == null) return acc;
						return acc.concat([v]);
					}, []);

					const aliasesNormalized = rawAliases.map((a) => (a ? normalizeKey(String(a)) : '')).filter(Boolean);
					return {
						name: sc.name || '',
						normalized: normalizeKey(sc.name || ''),
						aliases: rawAliases,
						aliasesNormalized,
						raw: sc,
					};
				});
				console.log('Settings identifier entries (update):', idInfo);
				notifyStatus(`Settings identifiers (update): ${idInfo.map((i) => i.name || '(unnamed)').join(', ') || '(none)'}`);
			} catch (logErr) {
				console.warn('Failed logging identifier entries from settings (update)', logErr);
			}
			// FIRST: detect identifier from the incoming file (try multiple candidates)
			const firstRow = dataToWrite[0];
			const isObjectRows = firstRow && typeof firstRow === 'object' && !Array.isArray(firstRow);

			// Build normalized header sets from incoming file/data
			const fileHeaderNames = isObjectRows ? Object.keys(firstRow || {}) : dataHeaders.slice();
			const fileHeaderKeySet = new Set((fileHeaderNames || []).map((h) => normalizeKey(String(h || ''))).filter(Boolean));
			const dataHeaderKeySet = new Set(Array.from(dataHeaderIndexMap.keys()));

			// First: find which settings entries are present in the incoming file (by name or aliases)
			const matchedSettingsInFile = [];
			for (let i = 0; i < settingsCols.length; i++) {
				const sc = settingsCols[i];
				if (!sc || !sc.name) continue;

				const nameKey = normalizeKey(String(sc.name));
				let matched = false;
				if (nameKey && (fileHeaderKeySet.has(nameKey) || dataHeaderKeySet.has(nameKey))) matched = true;

				// collect aliases (flatten)
				let aliases = [];
				if (Array.isArray(sc.aliases)) aliases = sc.aliases.slice();
				else if (Array.isArray(sc.alias)) aliases = sc.alias.slice();
				else if (typeof sc.alias === 'string') aliases = [sc.alias];
				if ((!aliases || aliases.length === 0) && Array.isArray(sc.alternates)) aliases = sc.alternates.slice();
				if ((!aliases || aliases.length === 0) && sc.alt) aliases = Array.isArray(sc.alt) ? sc.alt.slice() : [sc.alt];
				aliases = aliases.reduce((acc, v) => acc.concat(Array.isArray(v) ? v.filter(Boolean) : (v == null ? [] : [v])), []);

				// check aliases normalized
				if (!matched && aliases.length > 0) {
					for (let a = 0; a < aliases.length; a++) {
						const key = normalizeKey(String(aliases[a] || ''));
						if (!key) continue;
						const inFile = fileHeaderKeySet.has(key);
						const inData = dataHeaderKeySet.has(key);
						console.debug(`Settings match check name="${sc.name}" alias="${aliases[a]}" normalized="${key}" => inFile=${inFile}, inData=${inData}`);
						if (inFile || inData) {
							matched = true;
							break;
						}
					}
				}

				if (matched) matchedSettingsInFile.push(sc);
			}

			// Log what settings matched the incoming file (regardless of identifier flag)
			try {
				const matchedNames = matchedSettingsInFile.map((s) => s && s.name ? s.name : '(unnamed)');
				console.log('Settings matched in file (name or alias):', matchedNames);
				notifyStatus(`Settings matched in file: ${matchedNames.join(', ') || '(none)'}`);
			} catch (e) { /* ignore logging errors */ }

			// Now restrict to those explicitly marked identifier:true
			const fileCandidates = matchedSettingsInFile.filter((s) => s && s.identifier);
			if (fileCandidates.length === 0 && matchedSettingsInFile.length > 0) {
				// matched settings exist but none are flagged identifier
				const matchedNames = matchedSettingsInFile.map((s) => s && s.name ? s.name : '(unnamed)');
				console.warn('Matched settings were found but none are flagged identifier:true:', matchedNames);
				notifyStatus(`Matched settings found but none marked identifier: ${matchedNames.join(', ')}`);
			}
			// if still none, diagnostics will follow below as before

			if (fileCandidates.length === 0) {
				// Detailed diagnostics to help understand why no identifier was found
				try {
					const settingsNames = settingsCols.map((sc) => (sc && sc.name ? String(sc.name).trim() : ''));
					const settingsKeys = settingsCols.map((sc) => (sc && sc.name ? normalizeKey(sc.name) : ''));
					const fileHeaderNames = isObjectRows ? Object.keys(firstRow || {}) : dataHeaders.slice();
					const fileHeaderKeys = (fileHeaderNames || []).map((h) => (h == null ? '' : normalizeKey(String(h))));
					console.warn('Identifier not found in file. Diagnostics follow.');
					console.log('settings (names):', settingsNames);
					console.log('settings (normalized keys):', settingsKeys);
					console.log('file headers:', fileHeaderNames);
					console.log('file headers (normalized):', fileHeaderKeys);
					if (isObjectRows && dataToWrite && dataToWrite.length > 0) {
						console.log('sample first data row:', dataToWrite[0]);
					}
					notifyStatus('Identifier column not found in file; see console for diagnostics.');
				} catch (diagErr) {
					console.warn('Failed producing diagnostics for missing identifier', diagErr);
				}

				notifyComplete({ success: false, reason: 'identifier-not-in-file' });
				return;
			}

			// Log workbook settings for update flow
			try {
				const settingsNames = settingsCols.map((sc) => (sc && sc.name ? String(sc.name).trim() : ''));
				console.log('Workbook settings columns (update):', settingsNames);
				notifyStatus(`Workbook settings (update): ${settingsNames.join(', ') || '(none)'}`);
			} catch (logErr) {
				console.warn('Failed logging workbook settings for update', logErr);
			}

			// Log the identifier candidates discovered in the incoming file
			try {
				const candidateNames = fileCandidates.map((c) => (c && c.name ? String(c.name).trim() : ''));
				console.log('Identifier candidates found in file:', candidateNames);
				notifyStatus(`Identifier candidates in file: ${candidateNames.join(', ') || '(none)'}`);
			} catch (logErr) {
				console.warn('Failed logging file identifier candidates', logErr);
			}

			// alias / canonicalization like Refresh (keep in-memory only)
			const matchedByName = new Map();
			for (let i = 0; i < settingsCols.length; i++) {
				const sc = settingsCols[i];
				if (!sc || !sc.name) continue;
				const nameKey = normalizeKey(sc.name);
				if (sheetHeaderIndexMap.has(nameKey) && dataHeaderIndexMap.has(nameKey)) {
					matchedByName.set(nameKey, String(sc.name).trim());
				}
			}
			const remainingSettings = settingsCols.filter((sc) => sc && sc.name && !matchedByName.has(normalizeKey(sc.name)));
			const sheetAliasMap = buildAliasMap(remainingSettings, sheetHeaders);
			const dataAliasMap = buildAliasMap(remainingSettings, dataHeaders);

			let effectiveSheetHeaders = sheetHeaders.slice();
			if (matchedByName.size > 0) {
				effectiveSheetHeaders = effectiveSheetHeaders.map((h) => {
					const key = normalizeKey(h);
					return matchedByName.has(key) ? matchedByName.get(key) : h;
				});
			}
			if (sheetAliasMap && sheetAliasMap.size > 0) {
				effectiveSheetHeaders = renameHeaderArray(effectiveSheetHeaders, sheetAliasMap);
			}

			// normalize incoming data column names to canonical names
			const combinedDataAliasMap = new Map();
			for (const [k, v] of matchedByName.entries()) combinedDataAliasMap.set(k, v);
			if (dataAliasMap && dataAliasMap.size > 0) {
				for (const [k, v] of dataAliasMap.entries()) combinedDataAliasMap.set(k, v);
			}
			if (combinedDataAliasMap && combinedDataAliasMap.size > 0) {
				const firstRowInner = dataToWrite[0];
				if (firstRowInner && typeof firstRowInner === 'object' && !Array.isArray(firstRowInner)) {
					dataToWrite = renameObjectRows(dataToWrite, combinedDataAliasMap);
				} else if (Array.isArray(firstRowInner)) {
					dataToWrite = renameArrayRows(dataToWrite, combinedDataAliasMap);
				}
			}

			// SECOND: try each candidate (found in the file) against the worksheet until one matches
			let identifierSettingFromFile = null;
			let identifierIndex = undefined;
			for (let ci = 0; ci < fileCandidates.length; ci++) {
				const candidate = fileCandidates[ci];
				const info = computeIdentifierListFromValues(usedValues, effectiveSheetHeaders, candidate, usedFormulas);
				if (info && typeof info.identifierIndex === 'number') {
					identifierSettingFromFile = candidate;
					identifierIndex = info.identifierIndex;
					break;
				}
			}
			if (!identifierSettingFromFile || typeof identifierIndex !== 'number') {
				notifyStatus('None of the file-based identifier candidates were found on the sheet; update aborted.');
				notifyComplete({ success: false, reason: 'identifier-not-on-sheet', attempted: fileCandidates.map((c) => c.name) });
				return;
			}

			// Use 'matched' prop to determine which columns to update
			//			const providedHeaders = Array.isArray(matched) ? matched.map((h) => (h == null ? '' : String(h).trim())) : [];
			//			if (!providedHeaders || providedHeaders.length === 0) {
			//				notifyStatus('No update headers (matched) provided; update aborted.');
			//				notifyComplete({ success: false, reason: 'no-headers' });
			//				return;
			//			}

			// Trim incoming matched entries and canonicalize them using settingsColumns (aliases -> canonical name).
			let providedHeaders = Array.isArray(matched) ? matched.map((h) => (h == null ? '' : String(h).trim())) : [];
			if (!providedHeaders || providedHeaders.length === 0) {
				notifyStatus('No update headers (matched) provided.');
				notifyComplete({ success: false, reason: 'no-headers' });
				return;
			}

			// Build alias -> canonical name map from settingsCols
			const aliasToCanonical = new Map();
			for (let i = 0; i < settingsCols.length; i++) {
				const sc = settingsCols[i];
				if (!sc || !sc.name) continue;
				const canonical = String(sc.name).trim();
				const nameKey = normalizeKey(canonical);
				if (nameKey) aliasToCanonical.set(nameKey, canonical);

				// collect aliases/alternate names (flatten)
				let aliases = [];
				if (Array.isArray(sc.aliases)) aliases = sc.aliases.slice();
				else if (Array.isArray(sc.alias)) aliases = sc.alias.slice();
				else if (typeof sc.alias === 'string') aliases = [sc.alias];
				if ((!aliases || aliases.length === 0) && Array.isArray(sc.alternates)) aliases = sc.alternates.slice();
				if ((!aliases || aliases.length === 0) && sc.alt) aliases = Array.isArray(sc.alt) ? sc.alt.slice() : [sc.alt];
				aliases = aliases.reduce((acc, v) => acc.concat(Array.isArray(v) ? v.filter(Boolean) : (v == null ? [] : [v])), []);

				for (let a = 0; a < aliases.length; a++) {
					const key = normalizeKey(String(aliases[a] || ''));
					if (key) aliasToCanonical.set(key, canonical);
				}
			}

			// Map provided headers via aliasToCanonical and dedupe while preserving order
			{
				const seen = new Set();
				const out = [];
				for (let i = 0; i < providedHeaders.length; i++) {
					const ph = providedHeaders[i];
					const key = normalizeKey(ph);
					let mapped = ph;
					if (key && aliasToCanonical.has(key)) mapped = aliasToCanonical.get(key);
					const mk = normalizeKey(mapped);
					if (!mk) continue;
					if (seen.has(mk)) continue;
					seen.add(mk);
					out.push(mapped);
				}
				providedHeaders = out;
			}

			if (!providedHeaders || providedHeaders.length === 0) {
				notifyStatus('No update headers (matched) provided after canonicalization; update aborted.');
				notifyComplete({ success: false, reason: 'no-headers' });
				return;
			}

			// build normalized map of effective sheet headers -> index
			const effectiveHeaderIndexMap = new Map();
			for (let i = 0; i < effectiveSheetHeaders.length; i++) {
				const h = effectiveSheetHeaders[i];
				const key = normalizeKey(h);
				if (key) effectiveHeaderIndexMap.set(key, i);
			}

			// determine which provided headers actually map to sheet columns
			const updateColumns = [];
			for (let i = 0; i < providedHeaders.length; i++) {
				const ph = providedHeaders[i];
				const key = normalizeKey(ph);
				if (!key) continue;
				if (effectiveHeaderIndexMap.has(key)) {
					updateColumns.push({ name: effectiveSheetHeaders[effectiveHeaderIndexMap.get(key)], index: effectiveHeaderIndexMap.get(key) });
				} else {
					// header not found on sheet; inform once
					notifyStatus(`Update header not found on sheet: "${ph}" (skipped)`);
				}
			}
			if (updateColumns.length === 0) {
				notifyStatus('None of the provided headers matched sheet columns; update aborted.');
				notifyComplete({ success: false, reason: 'no-matching-headers' });
				return;
			}

			// build a mapping of existing ids -> absolute row index for efficient updates
			const existingIdToAbsRow = new Map();
			if (usedValues && usedValues.length > 0) {
				for (let r = 1; r < usedValues.length; r++) {
					const row = Array.isArray(usedValues[r]) ? usedValues[r] : [];
					let raw = row[identifierIndex];
					if (usedFormulas && usedFormulas[r] && typeof usedFormulas[r][identifierIndex] === 'string' && usedFormulas[r][identifierIndex].trim().startsWith('=')) {
						const link = extractHyperlink(usedFormulas[r][identifierIndex]);
						if (link) raw = link;
					}
					const id = raw == null ? '' : String(raw).trim();
					if (id) {
						const absRow = sheetRangeRowIndex + r;
						existingIdToAbsRow.set(id, absRow);
					}
				}
			}

			// Build updates/appends but populate only the provided updateColumns
			const updates = []; // { absRow, cells: [{colIdx, value}] }
			let skippedRows = 0; // count of incoming rows that would have been appended but are skipped for Update mode (no-append)

			const colCount = Math.max(sheetColCount, effectiveSheetHeaders.length, 1);

			for (let i = 0; i < dataToWrite.length; i++) {
				const rowIn = dataToWrite[i];
				// determine incoming id using the identifier found in the file
				let incomingId = '';
				if (isObjectRows && identifierSettingFromFile && identifierSettingFromFile.name) {
					const idKey = identifierSettingFromFile.name;
					const val = rowIn && rowIn[idKey];
					incomingId = val == null ? '' : String(val).trim();
				} else if (!isObjectRows) {
					if (identifierSettingFromFile && identifierSettingFromFile.name && dataHeaderIndexMap.has(normalizeKey(identifierSettingFromFile.name))) {
						const dataIdx = dataHeaderIndexMap.get(normalizeKey(identifierSettingFromFile.name));
						incomingId = (Array.isArray(rowIn) && rowIn.length > dataIdx) ? String(rowIn[dataIdx] ?? '').trim() : '';
					}
				}

				// Build cell updates for this incoming row
				const cellUpdates = [];
				for (let uc = 0; uc < updateColumns.length; uc++) {
					const col = updateColumns[uc];
					let val = '';
					if (isObjectRows) {
						val = rowIn && Object.prototype.hasOwnProperty.call(rowIn, col.name) ? rowIn[col.name] : '';
					} else {
						const key = normalizeKey(col.name);
						if (dataHeaderIndexMap.has(key)) {
							const dataIdx = dataHeaderIndexMap.get(key);
							val = (Array.isArray(rowIn) && rowIn.length > dataIdx) ? rowIn[dataIdx] : '';
						} else {
							val = '';
						}
					}
					cellUpdates.push({ colIdx: col.index, value: val });
				}

				// Only update existing rows. If identifier missing or not found, skip (no append).
				if (incomingId && existingIdToAbsRow.has(incomingId)) {
					updates.push({ absRow: existingIdToAbsRow.get(incomingId), cells: cellUpdates });
				} else {
					skippedRows++;
				}
			}

			// Perform writes in a single Excel.run: per-cell updates (queued) and a bulk append if needed.
			await Excel.run(async (context) => {
				const sheets = context.workbook.worksheets;
				const sheet = sheets.getItemOrNullObject(sheetName);
				await context.sync();
				if (sheet.isNullObject) {
					notifyStatus('Sheet disappeared before update could run.');
					notifyComplete({ success: false, reason: 'sheet-missing-during-update' });
					return;
				}

				// execute updates: write individual cells for specified columns
				for (let u = 0; u < updates.length; u++) {
					const upd = updates[u];
					for (let c = 0; c < upd.cells.length; c++) {
						const cell = upd.cells[c];
						try {
							const rng = sheet.getRangeByIndexes(upd.absRow, cell.colIdx, 1, 1);
							rng.values = [[cell.value]];
						} catch (e) {
							console.error('Update: failed writing cell', upd.absRow, cell.colIdx, e);
						}
					}
				}

				await context.sync();
			});

			notifyStatus(`Update completed (updated=${updates.length}, skipped=${skippedRows}).`);
			notifyComplete({ success: true, rowsUpdated: updates.length, rowsSkipped: skippedRows });
		} catch (err) {
			console.error('Update error', err);
			notifyStatus(`Update error: ${String(err)}`);
			notifyComplete({ success: false, error: String(err) });
		}
	}

	useEffect(() => {
		// Only call Refresh when explicitly requested via action === 'Refresh'.
		if (action === 'Refresh') {
			if (data && sheetName) {
				Refresh(data);
			} else {
				// missing params — report and complete
				if (!data) notifyStatus('Skipping refresh: no data provided.');
				if (!sheetName) notifyStatus('Skipping refresh: no sheetName provided.');
				notifyComplete({ success: false, reason: 'missing-params' });
			}
		} else if (action === 'Update') {
			// call Update when explicitly requested
			if (data && sheetName) {
				Update(data);
			} else {
				if (!data) notifyStatus('Skipping update: no data provided.');
				if (!sheetName) notifyStatus('Skipping update: no sheetName provided.');
				notifyComplete({ success: false, reason: 'missing-params' });
			}
		} else {
			// do nothing but return a status message per spec
			const actDesc = action === undefined || action === null ? 'no action provided' : `action="${String(action)}"`;
			notifyStatus(`Skipping refresh: ${actDesc}`);
			notifyComplete({ success: false, reason: 'action does not exist', action });
		}
	}, [data, sheetName, action]);

	return null;
}
