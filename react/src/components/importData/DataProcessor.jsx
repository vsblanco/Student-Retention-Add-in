import React, { useEffect, useState } from 'react';

export default function DataProcessor({ data, sheetName, matched, settingsColumns = [] }) {
	const [status, setStatus] = useState('');

	useEffect(() => {
		if (!data) return;
		if (!sheetName) {
			const msg = 'DataProcessor: sheetName is required to write data.';
			console.warn(msg);
			setStatus(msg);
			return;
		}
		// Log identifier columns (those with identifier === true)
		if (typeof DataProcessor.logIdentifierColumns === 'function') {
			DataProcessor.logIdentifierColumns(settingsColumns);
		}
		writeToWorksheet(data);
		// include settingsColumns & matched so changes update mapping/import behavior
		// eslint-disable-next-line react-hooks/exhaustive-deps
	}, [data, sheetName, settingsColumns, matched]);

	const to2D = (input) => {
		if (!Array.isArray(input) || input.length === 0) return [[]];
		const first = input[0];
		// Already rows of arrays
		if (Array.isArray(first)) return input;
		// Array of objects -> headers + rows
		if (first && typeof first === 'object') {
			const headers = Object.keys(first);
			const rows = input.map((r) => headers.map((k) => (r[k] !== undefined ? r[k] : '')));
			return [headers, ...rows];
		}
		// Fallback: single-column
		return input.map((v) => [v]);
	};

	// Normalize keys for matching: remove ALL whitespace and lowercase
	const normalizeKey = (val) => {
		if (val === null || val === undefined) return '';
		return String(val).replace(/\s/g, '').toLowerCase();
	};

	// Build alias -> name map from settingsColumns
	const buildAliasMap = (cols) => {
		const map = new Map();
		if (!Array.isArray(cols)) return map;
		cols.forEach((c) => {
			if (!c) return;
			if (typeof c === 'string') {
				// store normalized key -> original canonical name
				map.set(normalizeKey(c), c);
			} else if (typeof c === 'object') {
				const name = c.name ? String(c.name) : null;
				const alias = c.alias;
				if (name) {
					map.set(normalizeKey(name), name);
				}
				if (alias) {
					if (Array.isArray(alias)) {
						alias.forEach((a) => { if (a) map.set(normalizeKey(a), name || a); });
					} else {
						map.set(normalizeKey(alias), name || String(alias));
					}
				}
			}
		});
		return map;
	};

	const aliasMap = buildAliasMap(settingsColumns);

	// new helper: build a normalized matched set (canonicalized using aliasMap)
	const buildMatchedSet = (matchedArr = []) => {
		const set = new Set();
		if (!Array.isArray(matchedArr)) return set;
		matchedArr.forEach((m) => {
			if (m === null || typeof m === 'undefined') return;
			const raw = String(m);
			const key = normalizeKey(raw);
			const canonical = aliasMap.has(key) ? aliasMap.get(key) : raw;
			set.add(normalizeKey(canonical));
		});
		return set;
	};

	// Replace writeToWorksheet with smarter header-aware import
	const writeToWorksheet = async (inputData) => {
		setStatus(`Writing to worksheet "${sheetName}"...`);
		const tableData = to2D(inputData);

		try {
			if (window.Excel && typeof window.Excel.run === 'function') {
				await window.Excel.run(async (context) => {
					const wb = context.workbook;
					const sheets = wb.worksheets;
					const maybeSheet = sheets.getItemOrNullObject(sheetName);
					maybeSheet.load('name, id');
					await context.sync();

					let sheet;
					if (maybeSheet.isNullObject) {
						// create sheet and write full table (no existing header to match)
						sheet = sheets.add(sheetName);
						const rowCount = tableData.length;
						const colCount = tableData[0] ? tableData[0].length : 0;
						const range = sheet.getRangeByIndexes(0, 0, rowCount, colCount);
						range.values = tableData;
						range.format.autofitColumns();
						sheet.activate();
						await context.sync();
					} else {
						sheet = maybeSheet;

						// Get used range (or null) to detect existing header row and dimensions
						const used = sheet.getUsedRangeOrNullObject();
						used.load(['rowCount', 'columnCount']);
						await context.sync();

						// If the sheet has any used cells, read the first row as headerCandidates
						let sheetHeaders = [];
						// captured static values keyed by identifier (filled below if applicable)
						let capturedStaticMap = new Map();
						if (!used.isNullObject && used.columnCount > 0) {
							const headerRange = sheet.getRangeByIndexes(0, 0, 1, used.columnCount);
							headerRange.load('values');
							await context.sync();
							// keep original header values trimmed
							const originalHeaders = (headerRange.values && headerRange.values[0])
								? headerRange.values[0].map((h) => (h === null || h === undefined ? '' : String(h).trim()))
								: [];
							// normalize sheet headers by replacing any alias with canonical name
							sheetHeaders = originalHeaders.map((h) => {
								const key = normalizeKey(h);
								return aliasMap.has(key) ? aliasMap.get(key) : h;
							});
							// if any header changed due to alias replacement, update the worksheet header row
							const needsUpdate = originalHeaders.length !== sheetHeaders.length
								|| originalHeaders.some((h, i) => String((sheetHeaders[i] || '')).trim() !== String((h || '')).trim());
							if (needsUpdate) {
								try {
									headerRange.values = [sheetHeaders];
									await context.sync();
								} catch (e) {
									/* eslint-disable no-console */
									console.warn('Failed to update sheet headers with canonical names', e);
									/* eslint-enable no-console */
								}
							}

							// NEW CHECK:
							// If the sheet contains any static columns (per settingsColumns), require an identifier column on the sheet.
							// If identifier is missing, abort the import.
							(() => {
								// derive static names and ordered identifier candidates from settingsColumns
								const settingsStaticNames = [];
								const identifierCandidates = [];
								if (Array.isArray(settingsColumns)) {
									settingsColumns.forEach((c) => {
										if (!c || typeof c !== 'object') return;
										if (c.static) {
											const name = c.name || (Array.isArray(c.alias) ? c.alias[0] : c.alias) || null;
											if (name) {
												const canonical = aliasMap.has(normalizeKey(name)) ? aliasMap.get(normalizeKey(name)) : name;
												settingsStaticNames.push(canonical);
											}
										}
										if (c.identifier || c.identifer) identifierCandidates.push(c);
									});
								}

								// If no configured static columns, nothing to check
								if (settingsStaticNames.length === 0) return;

								// If none of the static columns are present on the sheet, nothing to check
								const staticOnSheet = sheetHeaders.some((h) =>
									settingsStaticNames.some((sn) => normalizeKey(h) === normalizeKey(sn))
								);
								if (!staticOnSheet) return;

								// If we have identifier candidates, prefer the first one that exists on the sheet.
								let chosenIdentifier = null;
								for (const cand of identifierCandidates) {
									const identifierName = cand && (cand.name || (Array.isArray(cand.alias) ? cand.alias[0] : cand.alias));
									const canonical = identifierName ? (aliasMap.has(normalizeKey(identifierName)) ? aliasMap.get(normalizeKey(identifierName)) : identifierName) : null;
									if (canonical && sheetHeaders.some((h) => normalizeKey(h) === normalizeKey(canonical))) {
										chosenIdentifier = cand;
										break;
									}
								}

								// A static column exists on the sheet — ensure an identifier candidate was found on the sheet
								if (!chosenIdentifier) {
									const first = identifierCandidates[0];
									const idDisplay = first && (first.name || (Array.isArray(first.alias) ? first.alias[0] : first.alias)) || 'identifier';
									const msg = `DataProcessor: static column(s) present on sheet but none of the configured identifier columns (first: "${idDisplay}") were found on sheet — import aborted.`;
									/* eslint-disable no-console */
									console.warn(msg);
									/* eslint-enable no-console */
									setStatus(msg);
									return;
								}
							})();

							// Capture existing static values keyed by identifier so we can restore them later.
							// This will only run when settingsColumns provide an identifier and at least one static column.
							try {
								capturedStaticMap = await DataProcessor.captureStaticColumns(settingsColumns, sheet, sheetHeaders, context, used);
							} catch (e) {
								/* eslint-disable no-console */
								console.warn('Failed to capture static column values', e);
								/* eslint-enable no-console */
								capturedStaticMap = new Map();
							}
						}

						// Determine whether input has headers (common case when to2D produced headers for array-of-objects)
						const inputHasHeaders = Array.isArray(tableData[0]) && tableData[0].every((v) => typeof v === 'string');
						let writeStartRow = 0;
						let valuesToWrite = [];
						let writeColCount = 0;

						if (sheetHeaders.length > 0) {
							// Sheet provides headers: map incoming columns by name (case-insensitive) and write under header row
							writeStartRow = 1; // do not overwrite existing header row
							writeColCount = sheetHeaders.length;

							// Build normalized matched set (if provided) to limit which columns we modify
							const matchedSet = buildMatchedSet(matched);

							// Determine which sheet columns are targeted by the matched set.
							// If matchedSet is empty, we assume all sheet columns are allowed to be written (legacy behavior).
							const targetedIndices = matchedSet.size > 0
								? sheetHeaders.map((sh, idx) => ({ sh, idx })).filter(({ sh }) => matchedSet.has(normalizeKey(sh))).map(({ idx }) => idx)
								: sheetHeaders.map((_, idx) => idx); // all indices

							if (inputHasHeaders) {
								let inputHeaders = tableData[0].map((h) => String(h).trim());
								// normalize input headers by alias -> name replacement
								inputHeaders = inputHeaders.map((h) => {
									const key = normalizeKey(h);
									return aliasMap.has(key) ? aliasMap.get(key) : h;
								});
								const dataRows = tableData.slice(1);

								// Build full-row arrays aligned to sheetHeaders (so restoreStaticColumns still works).
								// Fill only targetedIndices with values from input; non-targeted columns remain '' to preserve them.
								valuesToWrite = dataRows.map((row) => {
									const full = sheetHeaders.map(() => '');
									targetedIndices.forEach((colIdx) => {
										const sh = sheetHeaders[colIdx];
										const ihIdx = inputHeaders.findIndex((ih) => normalizeKey(ih) === normalizeKey(sh));
										full[colIdx] = ihIdx >= 0 ? (row[ihIdx] !== undefined ? row[ihIdx] : '') : '';
									});
									return full;
								});
							} else {
								// No input headers: align by input column order to targetedIndices order.
								// We'll map input column 0 -> targetedIndices[0], input 1 -> targetedIndices[1], etc.
								const dataRows = tableData;
								valuesToWrite = dataRows.map((row) => {
									const full = sheetHeaders.map(() => '');
									for (let i = 0; i < targetedIndices.length; i += 1) {
										const colIdx = targetedIndices[i];
										full[colIdx] = row[i] !== undefined ? row[i] : '';
									}
									return full;
								});
							}

							// Ensure at least one row/col to write (API doesn't like 0 dimensions)
							const rowCount = valuesToWrite.length || 1;
							writeColCount = writeColCount || (valuesToWrite[0] ? valuesToWrite[0].length : 0);

							// Clear previous data rows only for targeted columns (preserve others)
							if (!used.isNullObject && used.rowCount > writeStartRow) {
								const clearRowCount = Math.max((used.rowCount - writeStartRow), rowCount);
								// Clear per-targeted-column to avoid wiping untouched columns
								targetedIndices.forEach((colIdx) => {
									const clearRange = sheet.getRangeByIndexes(writeStartRow, colIdx, clearRowCount, 1);
									clearRange.clear(window.Excel ? window.Excel.ClearApplyTo.contents : 1);
								});
							}

							// Write targeted columns only (write per-column to preserve other columns)
							const rowCountToWrite = valuesToWrite.length || 1;
							// Prepare per-column arrays
							const perColumnWrites = targetedIndices.map((colIdx) => {
								const colVals = [];
								for (let r = 0; r < rowCountToWrite; r += 1) {
									const v = (valuesToWrite[r] && typeof valuesToWrite[r][colIdx] !== 'undefined') ? valuesToWrite[r][colIdx] : '';
									colVals.push([v]);
								}
								// assign to range immediately (defer sync until after loop)
								const writeRange = sheet.getRangeByIndexes(writeStartRow, colIdx, rowCountToWrite, 1);
								writeRange.values = colVals;
								return writeRange;
							});

							// Autofit columns for targeted columns
							targetedIndices.forEach((colIdx) => {
								try {
									const colRange = sheet.getRangeByIndexes(0, colIdx, Math.max(rowCountToWrite + 1, 2), 1);
									colRange.format.autofitColumns();
								} catch (e) {
									// ignore autofit errors
								}
							});

							sheet.activate();
							await context.sync();

							// After writing, restore any captured static values into their columns by matching identifier values.
							if (capturedStaticMap && capturedStaticMap.size > 0) {
								try {
									await DataProcessor.restoreStaticColumns(capturedStaticMap, settingsColumns, sheet, sheetHeaders, valuesToWrite, writeStartRow, context);
								} catch (e) {
									/* eslint-disable no-console */
									console.warn('Failed to restore static column values', e);
									/* eslint-enable no-console */
								}
							}
						} else {
							// No existing header on sheet: write full table starting at top-left
							writeStartRow = 0;
							valuesToWrite = tableData;
							writeColCount = tableData[0] ? tableData[0].length : 0;

							// Ensure at least one row/col to write (API doesn't like 0 dimensions)
							const rowCount = valuesToWrite.length || 1;
							writeColCount = writeColCount || (valuesToWrite[0] ? valuesToWrite[0].length : 0);

							// Clear previous data if any (full clear)
							if (!used.isNullObject && used.rowCount > 0) {
								const clearRowCount = Math.max(used.rowCount, rowCount);
								const clearRange = sheet.getRangeByIndexes(writeStartRow, 0, clearRowCount, writeColCount);
								clearRange.clear(window.Excel ? window.Excel.ClearApplyTo.contents : 1);
							}

							// Write full block
							const targetRange = sheet.getRangeByIndexes(writeStartRow, 0, rowCount, writeColCount);
							targetRange.values = valuesToWrite.length ? valuesToWrite : [['']];
							try {
								targetRange.format.autofitColumns();
							} catch (e) {
								// ignore
							}
							sheet.activate();
							await context.sync();
						}
					}
				});
				setStatus(`Written to worksheet "${sheetName}".`);
			} else {
				// Fallback when Office Excel API is not available in this environment
				console.log(`DataProcessor fallback output (no Excel API) for sheet "${sheetName}":`, tableData);
				setStatus('Excel API not available — data logged to console.');
			}
		} catch (error) {
			console.error(error);
			setStatus(`Failed to write to worksheet: ${error && error.message ? error.message : error}`);
		}
	};

	// attach a static helper to read settingsColumns and log names with static === true
	DataProcessor.logStaticColumns = function logStaticColumns(settingsColumns = []) {
		const statics = [];
		if (Array.isArray(settingsColumns)) {
			settingsColumns.forEach((c) => {
				if (!c) return;
				let name = null;
				let isStatic = false;
				if (typeof c === 'string') {
					name = c;
				} else if (typeof c === 'object') {
					// prefer explicit name, fall back to alias (first alias if array)
					name = c.name || (Array.isArray(c.alias) ? c.alias[0] : c.alias) || null;
					isStatic = !!c.static;
				}
				if (isStatic && name) statics.push(name);
			});
		}
		try {
			/* eslint-disable no-console */
			console.log('DataProcessor: static columns ->', statics);
			/* eslint-enable no-console */
		} catch (e) {
			// ignore logging errors
		}
		return statics;
	};

	// Capture static column values keyed by identifier from the existing worksheet data.
	// Returns a Map<identifierValueString, { [staticName]: value, ... }>
	DataProcessor.captureStaticColumns = async function captureStaticColumns(settingsColumns = [], sheet, sheetHeaders = [], context, used) {
		const result = new Map();
		if (!sheet || !context || !Array.isArray(sheetHeaders) || sheetHeaders.length === 0) return result;

		// Determine identifier name and static column names from settingsColumns
		const identifierCandidates = [];
		const staticNames = [];
		if (Array.isArray(settingsColumns)) {
			settingsColumns.forEach((c) => {
				if (!c || typeof c !== 'object') return;
				if (c.identifier || c.identifer) identifierCandidates.push(c);
				if (c.static) {
					const name = c.name || (Array.isArray(c.alias) ? c.alias[0] : c.alias) || null;
					if (name) staticNames.push(name);
				}
			});
		}
		if (identifierCandidates.length === 0 || staticNames.length === 0) return result;

		// prefer the first candidate that exists on the sheet
		let identifierEntry = null;
		for (const cand of identifierCandidates) {
			const identifierName = cand.name || (Array.isArray(cand.alias) ? cand.alias[0] : cand.alias);
			const canonical = identifierName ? (aliasMap.has(normalizeKey(identifierName)) ? aliasMap.get(normalizeKey(identifierName)) : identifierName) : null;
			if (canonical && sheetHeaders.some((h) => normalizeKey(h) === normalizeKey(canonical))) {
				identifierEntry = cand;
				break;
			}
		}
		if (!identifierEntry) return result;

		const identifierName = identifierEntry.name || (Array.isArray(identifierEntry.alias) ? identifierEntry.alias[0] : identifierEntry.alias);
		if (!identifierName) return result;

		// Map names to canonical names using aliasMap if available
		const canonicalIdentifier = aliasMap.has(normalizeKey(identifierName)) ? aliasMap.get(normalizeKey(identifierName)) : identifierName;
		const canonicalStatics = staticNames.map((n) => (aliasMap.has(normalizeKey(n)) ? aliasMap.get(normalizeKey(n)) : n));

		const idIndex = sheetHeaders.findIndex((h) => normalizeKey(h) === normalizeKey(canonicalIdentifier));
		const staticIndices = canonicalStatics.map((cn) => ({ name: cn, idx: sheetHeaders.findIndex((h) => normalizeKey(h) === normalizeKey(cn)) }))
			.filter((s) => s.idx >= 0);

		if (idIndex < 0 || staticIndices.length === 0) return result;

		const dataRowCount = Math.max((used && used.rowCount ? used.rowCount : 0) - 1, 0);
		if (dataRowCount <= 0) return result;

		const dataRange = sheet.getRangeByIndexes(1, 0, dataRowCount, Math.max(used.columnCount || 0, sheetHeaders.length));
		dataRange.load('values');
		await context.sync();

		(dataRange.values || []).forEach((row) => {
			const idVal = row[idIndex];
			if (idVal === null || typeof idVal === 'undefined' || String(idVal).trim() === '') return;
			const key = String(idVal);
			const entry = {};
			staticIndices.forEach((s) => {
				entry[s.name] = row[s.idx];
			});
			result.set(key, entry);
		});
		return result;
	};

	// Restore captured static column values into the sheet for the newly written rows.
	DataProcessor.restoreStaticColumns = async function restoreStaticColumns(staticMap = new Map(), settingsColumns = [], sheet, sheetHeaders = [], valuesToWrite = [], writeStartRow = 0, context) {
		if (!sheet || !context || !Array.isArray(sheetHeaders) || sheetHeaders.length === 0) return;
		if (!staticMap || staticMap.size === 0) return;

		// Determine identifier name from settings
		const identifierCandidates = [];
		const staticNames = [];
		if (Array.isArray(settingsColumns)) {
			settingsColumns.forEach((c) => {
				if (!c || typeof c !== 'object') return;
				if (c.identifier || c.identifer) identifierCandidates.push(c);
				if (c.static) {
					const name = c.name || (Array.isArray(c.alias) ? c.alias[0] : c.alias) || null;
					if (name) staticNames.push(name);
				}
			});
		}
		if (identifierCandidates.length === 0 || staticNames.length === 0) return;

		// prefer the first candidate that exists on the sheet
		let identifierEntry = null;
		for (const cand of identifierCandidates) {
			const identifierName = cand.name || (Array.isArray(cand.alias) ? cand.alias[0] : cand.alias);
			const canonical = identifierName ? (aliasMap.has(normalizeKey(identifierName)) ? aliasMap.get(normalizeKey(identifierName)) : identifierName) : null;
			if (canonical && sheetHeaders.some((h) => normalizeKey(h) === normalizeKey(canonical))) {
				identifierEntry = cand;
				break;
			}
		}
		if (!identifierEntry) return;

		const identifierName = identifierEntry.name || (Array.isArray(identifierEntry.alias) ? identifierEntry.alias[0] : identifierEntry.alias);
		const canonicalIdentifier = aliasMap.has(normalizeKey(identifierName)) ? aliasMap.get(normalizeKey(identifierName)) : identifierName;
		const idIndex = sheetHeaders.findIndex((h) => normalizeKey(h) === normalizeKey(canonicalIdentifier));
		if (idIndex < 0) return;

		// Canonical static names and their sheet indices
		const canonicalStatics = staticNames.map((n) => (aliasMap.has(normalizeKey(n)) ? aliasMap.get(normalizeKey(n)) : n));
		const staticIndices = canonicalStatics.map((cn) => ({ name: cn, idx: sheetHeaders.findIndex((h) => normalizeKey(h) === normalizeKey(cn)) }))
			.filter((s) => s.idx >= 0);
		if (staticIndices.length === 0) return;

		const rowCount = valuesToWrite.length || 0;
		if (rowCount === 0) return;

		// For each static column, prepare column values to write back based on the identifier mapping
		for (const s of staticIndices) {
			const colValues = [];
			for (let r = 0; r < rowCount; r += 1) {
				const idVal = valuesToWrite[r] && valuesToWrite[r][idIndex];
				const key = (idVal === null || typeof idVal === 'undefined') ? '' : String(idVal);
				let v = '';
				if (key && staticMap.has(key)) {
					v = staticMap.get(key)[s.name];
				}
				colValues.push([typeof v === 'undefined' ? '' : v]);
			}
			const writeRange = sheet.getRangeByIndexes(writeStartRow, s.idx, rowCount, 1);
			writeRange.values = colValues;
		}
		await context.sync();
	};

	// New: log columns where identifier === true (accepts either "identifier" or "identifer")
	DataProcessor.logIdentifierColumns = function logIdentifierColumns(settingsColumns = []) {
		const identifiers = [];
		if (Array.isArray(settingsColumns)) {
			settingsColumns.forEach((c) => {
				if (!c) return;
				let name = null;
				let isIdentifier = false;
				if (typeof c === 'string') {
					// string entries have no metadata -> skip
					name = c;
				} else if (typeof c === 'object') {
					name = c.name || (Array.isArray(c.alias) ? c.alias[0] : c.alias) || null;
					// accept both correct and misspelled property names
					isIdentifier = !!(c.identifier || c.identifer);
				}
				if (isIdentifier && name) identifiers.push(name);
			});
		}
		try {
			/* eslint-disable no-console */
			console.log('DataProcessor: identifier columns ->', identifiers);
			/* eslint-enable no-console */
		} catch (e) {
			// ignore logging errors
		}
		return identifiers;
	};

	return (
		<div style={{ marginTop: 8 }}>
			{/* minimal UI showing DataProcessor status */}
			{status && <div>DataProcessor: {status}</div>}
		</div>
	);
}
