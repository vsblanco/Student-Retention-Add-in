import React, { useEffect } from 'react';

export default function DataProcessor({ data, sheetName, settingsColumns, headers }) {
	// normalizeKey: trim, lowercase, and remove internal whitespace for robust matching
	function normalizeKey(v) {
		if (v == null) return '';
		return String(v).trim().toLowerCase().replace(/\s+/g, '');
	}

	// NEW helper: find column index in a header array for a settings entry (match name first, then aliases)
	function findColumnIndex(headerArr = [], setting = {}) {
		if (!Array.isArray(headerArr) || headerArr.length === 0 || !setting) return -1;
		const normalizedHeaders = headerArr.map(normalizeKey);
		// check canonical name first
		if (setting.name) {
			const nk = normalizeKey(setting.name);
			const idx = normalizedHeaders.indexOf(nk);
			if (idx !== -1) return idx;
		}
		// then check aliases
		if (setting.alias) {
			if (Array.isArray(setting.alias)) {
				for (const a of setting.alias) {
					const idx = normalizedHeaders.indexOf(normalizeKey(a));
					if (idx !== -1) return idx;
				}
			} else {
				const idx = normalizedHeaders.indexOf(normalizeKey(setting.alias));
				if (idx !== -1) return idx;
			}
		}
		return -1;
	}

	// checkAliases: match settingsColumns (name + alias) against targetColumns (worksheet or data headers)
	function checkAliases(settingsCols = [], targetCols = []) {
		// changed: use normalizeKey for consistent trimming + lowercase + whitespace removal
		const targetMap = new Map();
		(targetCols || []).forEach((c, idx) => {
			targetMap.set(normalizeKey(c), { name: c, index: idx });
		});

		const matches = (settingsCols || []).map((sc) => {
			if (!sc) return null;

			// 1) check the canonical name first
			if (sc.name) {
				const nameKey = normalizeKey(sc.name);
				if (targetMap.has(nameKey)) {
					const info = targetMap.get(nameKey);
					return { setting: sc, matchedName: info.name, matchedIndex: info.index };
				}
			}

			// 2) fall back to aliases if the name didn't match
			const candidates = [];
			if (sc.alias) {
				if (Array.isArray(sc.alias)) candidates.push(...sc.alias);
				else candidates.push(sc.alias);
			}
			const found = candidates.map(normalizeKey).find((n) => targetMap.has(n));
			if (found) {
				const info = targetMap.get(found);
				return { setting: sc, matchedName: info.name, matchedIndex: info.index };
			}

			// no match
			return { setting: sc, matchedName: null, matchedIndex: -1 };
		});

		const unmatched = matches.filter((m) => m && m.matchedName == null);
		return { matches, unmatched };
	}

	// buildAliasMap: returns a Map where normalized target column -> canonical setting.name
	function buildAliasMap(settingsCols = [], targetCols = []) {
		const targetSet = new Set((targetCols || []).map(normalizeKey));
		const map = new Map();

		(settingsCols || []).forEach((sc) => {
			if (!sc) return;
			const canonical = sc.name ? String(sc.name).trim() : null;
			if (!canonical) return;
			const candidates = new Set();
			candidates.add(normalizeKey(canonical));
			if (sc.alias) {
				if (Array.isArray(sc.alias)) sc.alias.forEach((a) => candidates.add(normalizeKey(a)));
				else candidates.add(normalizeKey(sc.alias));
			}
			// find which target column matches any candidate, map that target to canonical
			for (const cand of candidates) {
				if (targetSet.has(cand)) {
					// map the target (normalized) -> canonical (original casing from settings)
					map.set(cand, canonical);
				}
			}
		});

		return map;
	}

	// rename header array using aliasMap (map keys are normalized target -> canonical)
	function renameHeaderArray(headerArr = [], aliasMap = new Map()) {
		return (headerArr || []).map((h) => {
			const key = normalizeKey(h);
			return aliasMap.has(key) ? aliasMap.get(key) : h;
		});
	}

	// write a header row back to the worksheet so renamed/canonical headers are visible
	async function writeHeaderToWorksheet(headerArr = []) {
		if (!sheetName || !headerArr || headerArr.length === 0) return;
		if (!window.Excel || !window.Excel.run) {
			console.error('Excel JS API is not available in this context.');
			return;
		}
		try {
			await Excel.run(async (context) => {
				const sheets = context.workbook.worksheets;
				let sheet = sheets.getItemOrNullObject(sheetName);
				await context.sync();
				if (sheet.isNullObject) {
					// create sheet when missing
					sheet = sheets.add(sheetName);
				}
				// ensure the range for the header exists and set the values (single-row)
				const range = sheet.getRangeByIndexes(0, 0, 1, headerArr.length);
				range.values = [headerArr];
				range.format.autofitColumns();
				await context.sync();
			});
		} catch (err) {
			console.error('writeHeaderToWorksheet error', err);
		}
	}

	// NEW: read and retain values for static columns, keyed by identifier
	// signature changed: accept sheetHeaders so identifier discovery uses the worksheet header mapping
	async function retainStaticColumns(staticCols = [], identifierSetting = null, sheetHeaders = []) {
		if (!sheetName || !Array.isArray(staticCols) || staticCols.length === 0) return null;
		if (!identifierSetting || !identifierSetting.name) {
			console.warn('retainStaticColumns: no identifier setting provided; skipping static retention');
			return null;
		}
		if (!window.Excel || !window.Excel.run) {
			console.error('Excel JS API not available for retainStaticColumns');
			return null;
		}

		try {
			return await Excel.run(async (context) => {
				const sheets = context.workbook.worksheets;
				const sheet = sheets.getItemOrNullObject(sheetName);
				await context.sync();
				if (sheet.isNullObject) {
					// nothing to retain if sheet missing
					return null;
				}

				const used = sheet.getUsedRangeOrNullObject();
				used.load(['values', 'rowCount', 'columnCount']);
				await context.sync();

				if (used.isNullObject) return null;
				const values = used.values || [];
				if (!values || values.length === 0) return null;

				// use provided sheetHeaders (if given) or fallback to top row from used range
				const headerRow = (Array.isArray(sheetHeaders) && sheetHeaders.length > 0)
					? sheetHeaders.map((v) => (v == null ? '' : String(v).trim()))
					: (values[0].map((v) => (v == null ? '' : String(v).trim())));

				const normalizedHeader = headerRow.map(normalizeKey);

				// find identifier index using name + alias matching against the worksheet header
				const identifierIndex = findColumnIndex(headerRow, identifierSetting);
				if (identifierIndex === -1) {
					console.warn('retainStaticColumns: identifier column not found on worksheet');
					return null;
				}

				// find indices for static columns (using same find logic to respect aliases/canonical names)
				const staticIndices = staticCols.map((colName) => {
					const idxDirect = normalizedHeader.indexOf(normalizeKey(colName));
					return idxDirect; // -1 if not present
				});

				// build map: identifierValue(string) -> { colName: value, ... }
				const saved = new Map();
				// iterate data rows (starting at row 1, since 0 is header)
				for (let r = 1; r < values.length; r++) {
					const row = Array.isArray(values[r]) ? values[r] : [];
					const idValRaw = row[identifierIndex];
					const idVal = idValRaw == null ? '' : String(idValRaw).trim();
					const obj = {};

					staticCols.forEach((colName, i) => {
						const colIdx = staticIndices[i];
						const rawVal = colIdx >= 0 && row.length > colIdx ? row[colIdx] : null;
						// only store when the cell has a non-empty value (null/undefined/empty-string => skip)
						if (rawVal != null) {
							const asStr = String(rawVal).trim();
							if (asStr !== '') {
								obj[colName] = rawVal;
							}
						}
					});

					// only add map entry when at least one static value was captured
					if (Object.keys(obj).length > 0) {
						saved.set(String(idVal), obj);
					}
				}

				// LOG: report static capture summary
				if (saved.size > 0) {
					console.log(`retainStaticColumns: captured ${saved.size} static row entries for cols: ${staticCols.join(', ')}`);
				} else {
					console.log('retainStaticColumns: no static values captured');
				}

				return {
					identifierIndex,
					staticCols,
					staticIndices,
					savedMap: saved,
				};
			});
		} catch (err) {
			console.error('retainStaticColumns error', err);
			return null;
		}
	}

	// NEW: restore static column values using the saved map, matching by identifier for the newly written rows
	async function restoreStaticColumns(savedInfo = null, writeStartRow = 0, rowCount = 0) {
		if (!savedInfo || !savedInfo.savedMap || savedInfo.savedMap.size === 0) return;
		if (!sheetName) return;
		if (!window.Excel || !window.Excel.run) {
			console.error('Excel JS API not available for restoreStaticColumns');
			return;
		}

		try {
			await Excel.run(async (context) => {
				const sheets = context.workbook.worksheets;
				const sheet = sheets.getItemOrNullObject(sheetName);
				await context.sync();
				if (sheet.isNullObject) return;

				// validate rowCount
				if (rowCount <= 0) return;

				// read identifiers for the newly written rows
				const idRange = sheet.getRangeByIndexes(writeStartRow, savedInfo.identifierIndex, rowCount, 1);
				idRange.load('values');
				await context.sync();

				const idValues = idRange.values || [];

				// for each static column, build values to write back
				for (let i = 0; i < savedInfo.staticCols.length; i++) {
					const colName = savedInfo.staticCols[i];
					const colIndex = savedInfo.staticIndices[i];
					if (colIndex === -1) continue; // skip missing columns

					// prepare a column of values (rowCount x 1)
					const valuesToWrite = idValues.map((row) => {
						const rawId = row && row[0] != null ? String(row[0]).trim() : '';
						const savedRow = savedInfo.savedMap.get(String(rawId));
						return [savedRow && Object.prototype.hasOwnProperty.call(savedRow, colName) ? savedRow[colName] : ''];
					});

					// write the column back
					const targetRange = sheet.getRangeByIndexes(writeStartRow, colIndex, rowCount, 1);
					targetRange.values = valuesToWrite;
				}

				await context.sync();
			});
		} catch (err) {
			console.error('restoreStaticColumns error', err);
		}
	}

	// gatherIdentifierColumn: find settingsColumns entry where identifier === true
	function gatherIdentifierColumn(settingsCols = []) {
		const cols = Array.isArray(settingsCols) ? settingsCols : (Array.isArray(settingsColumns) ? settingsColumns : []);
		if (!cols || cols.length === 0) {
			console.log('gatherIdentifierColumn: no settings columns provided');
			return null;
		}
		// support boolean true, string "true", and common misspelling 'identifer'
		const found = cols.find((c) => {
			if (!c || typeof c !== 'object') return false;
			const id = c.identifier;
			const idMiss = c.identifer; // tolerate misspelling seen elsewhere
			return id === true || id === 'true' || idMiss === true || idMiss === 'true';
		}) || null;
		console.log('gatherIdentifierColumn found:', found);
		return found;
	}

	// rename object keys for array of objects using aliasMap
	function renameObjectRows(rows = [], aliasMap = new Map()) {
		if (!aliasMap || aliasMap.size === 0) return rows.slice();
		return rows.map((obj) => {
			if (!obj || typeof obj !== 'object' || Array.isArray(obj)) return obj;
			const newObj = {};
			Object.keys(obj).forEach((k) => {
				const nk = normalizeKey(k);
				const canonical = aliasMap.get(nk);
				if (canonical) newObj[canonical] = obj[k];
				else newObj[k] = obj[k];
			});
			return newObj;
		});
	}

	// rename first-row header of array-of-arrays and keep remaining rows unchanged
	function renameArrayRows(rows = [], aliasMap = new Map()) {
		if (!Array.isArray(rows) || rows.length === 0 || aliasMap.size === 0) return rows.slice();
		const header = rows[0];
		const renamedHeader = renameHeaderArray(header, aliasMap);
		return [renamedHeader, ...rows.slice(1)];
	}

	// gatherWorksheetColumns: return header row (array of strings; empty strings preserved)
	async function gatherWorksheetColumns() {
		if (!sheetName) return [];
		if (!window.Excel || !window.Excel.run) {
			console.error('Excel JS API is not available in this context.');
			return [];
		}

		try {
			return await Excel.run(async (context) => {
				const sheets = context.workbook.worksheets;
				const sheet = sheets.getItemOrNullObject(sheetName);
				await context.sync();
				if (sheet.isNullObject) return [];

				const used = sheet.getUsedRangeOrNullObject();
				used.load('values');
				await context.sync();

				if (used.isNullObject) return [];

				const raw = used.values && used.values[0] ? used.values[0] : [];
				return raw.map((v) => (v == null ? '' : String(v).trim()));
			});
		} catch (err) {
			console.error('gatherWorksheetColumns error', err);
			return [];
		}
	}

	// NEW helper: compute saved static data from a used-range values array (no Excel.run)
	function computeSavedStaticFromValues(values = [], sheetHeaders = [], staticCols = [], identifierSetting = null) {
		if (!Array.isArray(staticCols) || staticCols.length === 0) return null;
		if (!identifierSetting || !identifierSetting.name) return null;
		if (!Array.isArray(values) || values.length === 0) return null;

		const headerRow = (Array.isArray(sheetHeaders) && sheetHeaders.length > 0)
			? sheetHeaders.map((v) => (v == null ? '' : String(v).trim()))
			: (values[0].map((v) => (v == null ? '' : String(v).trim())));

		const normalizedHeader = headerRow.map(normalizeKey);

		const identifierIndex = findColumnIndex(headerRow, identifierSetting);
		if (identifierIndex === -1) return null;

		const staticIndices = staticCols.map((colName) => normalizedHeader.indexOf(normalizeKey(colName)));

		const saved = new Map();
		for (let r = 1; r < values.length; r++) {
			const row = Array.isArray(values[r]) ? values[r] : [];
			const idValRaw = row[identifierIndex];
			const idVal = idValRaw == null ? '' : String(idValRaw).trim();
			const obj = {};

			staticCols.forEach((colName, i) => {
				const colIdx = staticIndices[i];
				const rawVal = colIdx >= 0 && row.length > colIdx ? row[colIdx] : null;
				// only keep non-empty values (preserve 0 and false)
				if (rawVal != null) {
					const asStr = String(rawVal).trim();
					if (asStr !== '') {
						obj[colName] = rawVal;
					}
				}
			});

			if (Object.keys(obj).length > 0) {
				saved.set(String(idVal), obj);
			}
		}

		// LOG: report computeSavedStaticFromValues summary
		if (saved.size > 0) {
			console.log(`computeSavedStaticFromValues: captured ${saved.size} static row entries for cols: ${staticCols.join(', ')}`);
		} else {
			console.log('computeSavedStaticFromValues: no static values captured');
		}

		return {
			identifierIndex,
			staticCols,
			staticIndices,
			savedMap: saved,
		};
	}

	// NEW helper: apply static columns inside an existing Excel.run context and sheet
	async function applyStaticColumnsWithContext(context, sheet, savedInfo, writeStartRow = 0, rowCount = 0) {
		if (!savedInfo || !savedInfo.savedMap || savedInfo.savedMap.size === 0) return;
		if (!sheet || rowCount <= 0) return;

		// read identifiers for the newly written rows
		const idRange = sheet.getRangeByIndexes(writeStartRow, savedInfo.identifierIndex, rowCount, 1);
		idRange.load('values');
		await context.sync();

		const idValues = idRange.values || [];

		for (let i = 0; i < savedInfo.staticCols.length; i++) {
			const colName = savedInfo.staticCols[i];
			const colIndex = savedInfo.staticIndices[i];
			if (colIndex === -1) continue;

			const valuesToWrite = idValues.map((row) => {
				const rawId = row && row[0] != null ? String(row[0]).trim() : '';
				const savedRow = savedInfo.savedMap.get(String(rawId));
				return [savedRow && Object.prototype.hasOwnProperty.call(savedRow, colName) ? savedRow[colName] : ''];
			});

			const targetRange = sheet.getRangeByIndexes(writeStartRow, colIndex, rowCount, 1);
			targetRange.values = valuesToWrite;

			// LOG: report each static column write
			console.log(`applyStaticColumnsWithContext: queued restore for static column '${colName}' (${valuesToWrite.length} rows) on sheet ${sheetName}`);
		}
	}

	// Refresh: write `data` to the worksheet named `sheetName` using the Excel JS API.
	async function Refresh(dataToWrite) {
		if (!dataToWrite || !dataToWrite.length) {
			console.warn('Refresh: no data to write');
			return;
		}
		if (!sheetName) {
			console.warn('Refresh: missing sheetName');
			return;
		}
		if (!window.Excel || !window.Excel.run) {
			console.error('Excel JS API is not available in this context.');
			return;
		}

		try {
			// INITIAL READ: get sheet (if exists) and used range values in one Excel.run
			let usedValues = [];
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
					used.load(['values']);
					await context.sync();
					if (!used.isNullObject && Array.isArray(used.values)) {
						usedValues = used.values;
					} else {
						usedValues = [];
					}
				}
			});

			// derive sheetHeaders from usedValues (if any)
			let sheetHeaders = Array.isArray(usedValues) && usedValues[0] ? usedValues[0].map((v) => (v == null ? '' : String(v).trim())) : [];
			const hasHeaders = Array.isArray(sheetHeaders) && sheetHeaders.some((h) => h !== '');

			// derive data headers from incoming data (object keys or first-row array)
			const dataFirst = dataToWrite[0];
			let dataHeaders = (dataFirst && typeof dataFirst === 'object' && !Array.isArray(dataFirst))
				? Object.keys(dataFirst).map((h) => (h == null ? '' : String(h).trim()))
				: (Array.isArray(dataFirst) ? dataFirst.map((h) => (h == null ? '' : String(h).trim())) : []);

			// normalize settings
			const settingsCols = Array.isArray(settingsColumns) ? settingsColumns : [];

			// 1) Check "name" matches first (exact canonical name present in both)
			const sheetSet = new Set((sheetHeaders || []).map(normalizeKey));
			const dataSet = new Set((dataHeaders || []).map(normalizeKey));
			const matchedByName = new Map();
			(settingsCols || []).forEach((sc) => {
				if (!sc || !sc.name) return;
				const nameKey = normalizeKey(sc.name);
				if (sheetSet.has(nameKey) && dataSet.has(nameKey)) {
					matchedByName.set(nameKey, String(sc.name).trim());
				}
			});

			// 2) For any settings not matched by name, run alias checks and rename where possible.
			const remainingSettings = (settingsCols || []).filter((sc) => {
				if (!sc || !sc.name) return false;
				return !matchedByName.has(normalizeKey(sc.name));
			});

			// build alias maps against worksheet and data (only for remaining settings)
			const sheetAliasMap = buildAliasMap(remainingSettings, sheetHeaders); // normalized target -> canonical
			const dataAliasMap = buildAliasMap(remainingSettings, dataHeaders);

			// 3) Apply canonicalization for exact-name matches on worksheet headers (so casing becomes settings.name)
			if (matchedByName.size > 0) {
				sheetHeaders = sheetHeaders.map((h) => {
					const key = normalizeKey(h);
					return matchedByName.has(key) ? matchedByName.get(key) : h;
				});
				// Note: header writes deferred to the single write Excel.run below
			}

			// 4) Apply alias renames to worksheet headers in-memory (write deferred)
			if (sheetAliasMap && sheetAliasMap.size > 0) {
				sheetHeaders = renameHeaderArray(sheetHeaders, sheetAliasMap);
				// write deferred
			}

			// 5) Normalize incoming data column names to canonical names (both name-matches and alias-matches)
			const combinedDataAliasMap = new Map();
			for (const [k, v] of matchedByName.entries()) combinedDataAliasMap.set(k, v);
			if (dataAliasMap && dataAliasMap.size > 0) {
				for (const [k, v] of dataAliasMap.entries()) combinedDataAliasMap.set(k, v);
			}

			let normalizedData = dataToWrite;
			if (combinedDataAliasMap && combinedDataAliasMap.size > 0) {
				const firstRow = dataToWrite[0];
				if (firstRow && typeof firstRow === 'object' && !Array.isArray(firstRow)) {
					normalizedData = renameObjectRows(dataToWrite, combinedDataAliasMap);
				} else if (Array.isArray(firstRow)) {
					normalizedData = renameArrayRows(dataToWrite, combinedDataAliasMap);
				}
			}
			dataToWrite = normalizedData;

			// --- determine static columns (present on worksheet but not in data) ---
			// Recompute the set of data fields AFTER normalization.
			// For object rows, use the union of keys across all objects (so any mapped key is considered non-static).
			// For array rows, use the first row as header positions.
			let dataFieldSet;
			const firstAfter = dataToWrite && dataToWrite[0];
			const isObjectRowsAfter = firstAfter && typeof firstAfter === 'object' && !Array.isArray(firstAfter);
			if (isObjectRowsAfter) {
				dataFieldSet = new Set();
				for (const rowObj of dataToWrite) {
					if (rowObj && typeof rowObj === 'object' && !Array.isArray(rowObj)) {
						Object.keys(rowObj).forEach((k) => {
							const key = normalizeKey(k);
							if (key) dataFieldSet.add(key);
						});
					}
				}
			} else if (Array.isArray(firstAfter)) {
				dataFieldSet = new Set((firstAfter || []).map((h) => normalizeKey(h)));
			} else {
				dataFieldSet = new Set();
			}

			const sheetHeadersNormalized = (sheetHeaders || []).map((v) => (v == null ? '' : String(v).trim()));
			const staticCols = sheetHeadersNormalized.filter((h) => {
				const key = normalizeKey(h);
				return key && !dataFieldSet.has(key);
			});

			// gather identifier setting and compute saved static values from the usedValues read above
			let savedStatic = null;
			if (staticCols.length > 0) {
				const identifierSetting = gatherIdentifierColumn(settingsCols);
				if (!identifierSetting || !identifierSetting.name) {
					console.warn('No identifier configured; skipping static column retention.');
				} else {
					// compute saved static from the values we already read (no Excel.run)
					savedStatic = computeSavedStaticFromValues(usedValues, sheetHeaders, staticCols, identifierSetting);
				}
			}

			// Build rows to write and decide writeStartRow
			const first = dataToWrite[0];
			const isObjectRows = first && typeof first === 'object' && !Array.isArray(first);

			let rows;
			let writeStartRow = 0;

			if (hasHeaders) {
				const desiredColumns = sheetHeaders
					.map((h, idx) => (h ? { name: h, index: idx } : null))
					.filter(Boolean);

				if (isObjectRows) {
					rows = dataToWrite.map((obj) =>
						desiredColumns.map((c) => (obj && Object.prototype.hasOwnProperty.call(obj, c.name) ? obj[c.name] : ''))
					);
				} else {
					rows = dataToWrite.map((arr) =>
						desiredColumns.map((c) => (Array.isArray(arr) && arr.length > c.index ? arr[c.index] : ''))
					);
				}

				writeStartRow = 1;
			} else {
				if (isObjectRows) {
					const headers = Object.keys(first);
					rows = [headers, ...dataToWrite.map((obj) => headers.map((h) => (obj[h] ?? '')))];
					writeStartRow = 0;
				} else {
					rows = dataToWrite;
					writeStartRow = 0;
				}
			}

			const rowCount = rows.length;
			let colCount = rows[0] ? rows[0].length : 0;
			if (rowCount === 0 || colCount === 0) {
				console.warn('Refresh: nothing to write after normalization');
				return;
			}

			rows = rows.map((r) => {
				const rowArr = Array.isArray(r) ? r.slice() : [];
				if (rowArr.length < colCount) return rowArr.concat(Array(colCount - rowArr.length).fill(''));
				if (rowArr.length > colCount) return rowArr.slice(0, colCount);
				return rowArr;
			});

			// FINAL WRITE: single Excel.run to create sheet (if needed), write header + data, and restore static columns
			await Excel.run(async (context) => {
				const sheets = context.workbook.worksheets;
				let sheet = sheets.getItemOrNullObject(sheetName);
				await context.sync();
				if (sheet.isNullObject) {
					sheet = sheets.add(sheetName);
				}

				// --- NEW: clear all existing rows below the header to remove leftover values before writing ---
				const used = sheet.getUsedRangeOrNullObject();
				used.load(['rowCount', 'columnCount']);
				await context.sync();

				if (!used.isNullObject) {
					const existingRowCount = used.rowCount || 0;
					const existingColCount = used.columnCount || 0;
					// only clear rows below the header (preserve header at row 0)
					const rowsToClear = Math.max(0, existingRowCount - writeStartRow);
					const colsToClear = Math.max(colCount, existingColCount, 1);
					if (rowsToClear > 0 && colsToClear > 0) {
						// build a blank matrix rowsToClear x colsToClear
						const blankRow = new Array(colsToClear).fill('');
						const blankMatrix = new Array(rowsToClear).fill(null).map(() => blankRow.slice());
						const clearRange = sheet.getRangeByIndexes(writeStartRow, 0, rowsToClear, colsToClear);
						clearRange.values = blankMatrix;

						// LOG: clearing summary
						console.log(`Cleared ${rowsToClear} rows and ${colsToClear} columns below header on sheet ${sheetName}`);
					}
				}

				// write header row if we have headers (use current sheetHeaders array)
				if (hasHeaders && Array.isArray(sheetHeaders) && sheetHeaders.length > 0) {
					const headerRange = sheet.getRangeByIndexes(0, 0, 1, sheetHeaders.length);
					headerRange.values = [sheetHeaders];

					// LOG: header written
					console.log(`Wrote header to sheet ${sheetName}`);
				}

				// write all rows in one range write
				const writeRange = sheet.getRangeByIndexes(writeStartRow, 0, rowCount, colCount);
				writeRange.values = rows;

				// LOG: paste/write summary
				console.log(`Pasted ${rowCount} rows x ${colCount} cols to sheet ${sheetName} starting at row ${writeStartRow}`);

				writeRange.format.autofitColumns();
				writeRange.format.autofitRows();

				// if we saved static columns, restore them here (reads ids then writes values)
				if (savedStatic && savedStatic.savedMap && savedStatic.savedMap.size > 0) {
					console.log(`Restoring static columns (${savedStatic.staticCols.join(', ')}) for ${savedStatic.savedMap.size} saved rows`);
					await applyStaticColumnsWithContext(context, sheet, savedStatic, writeStartRow, rowCount);
				}

				await context.sync();
			});
		} catch (err) {
			console.error('Refresh error', err);
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
