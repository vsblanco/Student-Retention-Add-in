import React, { useState, useRef, useMemo, useEffect } from 'react';
import parseCSV from './Parsers/csv';
import DataProcessor from './DataProcessor';
import styles from './importManagerStyles';
import { getImportType } from './ImportType';
import { getWorkbookSettings } from '../utility/getSettings';
import { Upload } from 'lucide-react';
import FileCard from './FileCard';
import ImportIcon from '../../assets/icons/import-icon.png';

export default function ImportManager({ onImport, excludeFilter, hyperLink } = {}) {
	const [fileName, setFileName] = useState(''); // name of active file
	const [status, setStatus] = useState('');
	const [parsedData, setParsedData] = useState(null);
	const [headers, setHeaders] = useState([]); // store column headers
	// now support multiple uploaded files
	const [uploadedFiles, setUploadedFiles] = useState([]); // array of File
	const [activeIndex, setActiveIndex] = useState(-1); // index into uploadedFiles that is parsed/active
	// per-file import info computed from headers (e.g. { type, action, matched })
	const [fileInfos, setFileInfos] = useState([]);
	const [isImported, setIsImported] = useState(false); // whether user clicked Import
	const [workbookColumns, setWorkbookColumns] = useState([]); // new: columns from workbook settings

	// NEW: index of file currently being processed (sequenced import)
	const [processingIndex, setProcessingIndex] = useState(-1);
	// NEW: track last index for which onImport was emitted so we don't emit twice
	const [lastEmittedIndex, setLastEmittedIndex] = useState(-1);

	const inputRef = useRef(null);

	// drag state for drop-zone
	const [dragActive, setDragActive] = useState(false);

	// parse a single CSV file into parsedData/headers and set active
	const parseCSVFile = (file, index) => {
		if (!file) return;
		setFileName(file.name);
		setStatus('Reading file...');
		// DO NOT clear isImported here — parsing is used during multi-file import as well
		// setIsImported(false);

		const isCSV = /\.csv$/i.test(file.name);
		if (!isCSV) {
			setStatus('Unsupported file type. Please select a .csv file.');
			return;
		}

		const reader = new FileReader();
		reader.onerror = () => setStatus('Failed to read file.');
		reader.onload = (e) => {
			try {
				const text = e.target.result;
				const data = parseCSV(text);

				// extract headers
				let extractedHeaders = [];
				if (Array.isArray(data) && data.length > 0) {
					const firstRow = data[0];
					if (firstRow && typeof firstRow === 'object' && !Array.isArray(firstRow)) {
						extractedHeaders = Object.keys(firstRow);
					} else if (Array.isArray(firstRow)) {
						extractedHeaders = firstRow;
					}
				}
				setHeaders(extractedHeaders);

				// call workbook settings util
				try {
					const wbSettings = getWorkbookSettings(extractedHeaders);
					setWorkbookColumns(Array.isArray(wbSettings.columns) ? [...wbSettings.columns] : []);
					if (typeof DataProcessor.logStaticColumns === 'function') {
						DataProcessor.logStaticColumns(wbSettings.columns);
					}
					/* eslint-disable no-console */
					console.log('ImportManager: workbook settings ->', wbSettings);
					/* eslint-enable no-console */
				} catch (err) {
					// ignore
				}

				setStatus(`Parsed ${Array.isArray(data) ? data.length : 0} rows`);
				setParsedData(data);
				setActiveIndex(index);
				// compute and store import info for this file
				try {
					const info = getImportType(extractedHeaders);
					setFileInfos((prev) => {
						const copy = [...prev];
						copy[index] = info;
						return copy;
					});
				} catch (e) {
					// ignore
				}
			} catch (err) {
				setStatus('Error parsing CSV.');
				console.error(err);
			}
		};
		reader.readAsText(file);
	};

	// add one or more files to the uploadedFiles array; parse the first CSV among the added files
	const handleAddFiles = (filesList) => {
		if (!filesList || filesList.length === 0) return;
		const filesArr = Array.from(filesList);

		// If we've already imported and the user selects/upload new files,
		// reset the uploadedFiles list and treat the incoming set as the next batch.
		if (isImported) {
			// dedupe incoming files by name (keep first occurrence)
			const seen = new Set();
			const uniques = [];
			filesArr.forEach((f) => {
				if (f && !seen.has(f.name)) {
					uniques.push(f);
					seen.add(f.name);
				}
			});

			// reset state to the new batch
			setUploadedFiles(uniques);
			setFileInfos(uniques.map(() => null));
			setParsedData(null);
			setHeaders([]);
			setActiveIndex(-1);
			setIsImported(false);

			// compute fileInfos for each new file by reading a small slice
			uniques.forEach((f, i) => {
				if (/\.csv$/i.test(f.name)) {
					const r = new FileReader();
					r.onload = (ev) => {
						try {
							const text = ev.target.result;
							const data = parseCSV(text);
							let extractedHeaders = [];
							if (Array.isArray(data) && data.length > 0) {
								const firstRow = data[0];
								if (firstRow && typeof firstRow === 'object' && !Array.isArray(firstRow)) {
									extractedHeaders = Object.keys(firstRow);
								} else if (Array.isArray(firstRow)) {
									extractedHeaders = firstRow;
								}
							}
							const info = getImportType(extractedHeaders);
							setFileInfos((prev) => {
								const copy = Array.isArray(prev) ? [...prev] : [];
								copy[i] = info;
								return copy;
							});
						} catch (err) {
							// ignore parse errors for the chunk
						}
					};
					r.onerror = () => { /* ignore */ };
					r.readAsText(f.slice(0, 64 * 1024));
				}
			});

			// parse the first CSV among the new batch and make it active
			const firstCsvIdx = uniques.findIndex((f) => /\.csv$/i.test(f.name));
			if (firstCsvIdx !== -1) {
				parseCSVFile(uniques[firstCsvIdx], firstCsvIdx);
			}
			return;
		}

		// existing behavior: append or replace files and expand/adjust fileInfos placeholders, then read a small slice of each CSV
		setUploadedFiles((prev) => {
			const copy = [...prev];
			// track whether we've parsed a new CSV among newly appended files (back-compat)
			let parsedNewCsv = false;

			filesArr.forEach((f) => {
				// find existing file with same name
				const existingIndex = copy.findIndex((p) => p.name === f.name);
				if (existingIndex !== -1) {
					// replace existing file
					copy[existingIndex] = f;
					// ensure fileInfos placeholder exists at that index
					setFileInfos((prevInfos) => {
						const infos = Array.isArray(prevInfos) ? [...prevInfos] : [];
						infos[existingIndex] = null;
						return infos;
					});
					// if CSV, read a small slice to recompute headers/importInfo and also fully parse to make active
					if (/\.csv$/i.test(f.name)) {
						const r = new FileReader();
						r.onload = (ev) => {
							try {
								const text = ev.target.result;
								const data = parseCSV(text);
								let extractedHeaders = [];
								if (Array.isArray(data) && data.length > 0) {
									const firstRow = data[0];
									if (firstRow && typeof firstRow === 'object' && !Array.isArray(firstRow)) {
										extractedHeaders = Object.keys(firstRow);
									} else if (Array.isArray(firstRow)) {
										extractedHeaders = firstRow;
									}
								}
								try {
									const info = getImportType(extractedHeaders);
									setFileInfos((prev2) => {
										const copy2 = Array.isArray(prev2) ? [...prev2] : [];
										copy2[existingIndex] = info;
										return copy2;
									});
								} catch (err) { /* ignore */ }
								// fully parse and make this replaced file active (preserve prior behavior for replacements)
								parseCSVFile(f, existingIndex);
							} catch (err) { /* ignore parse errors for the chunk */ }
						};
						r.onerror = () => { /* ignore */ };
						r.readAsText(f.slice(0, 64 * 1024));
					}
				} else {
					// new file: append
					const newIndex = copy.length;
					copy.push(f);
					// expand fileInfos
					setFileInfos((prevInfos) => [...(Array.isArray(prevInfos) ? prevInfos : []), null]);

					// for each new unique file, try to read a small slice to detect headers and compute importInfo
					if (/\.csv$/i.test(f.name)) {
						const r = new FileReader();
						r.onload = (ev) => {
							try {
								const text = ev.target.result;
								const data = parseCSV(text);
								let extractedHeaders = [];
								if (Array.isArray(data) && data.length > 0) {
									const firstRow = data[0];
									if (firstRow && typeof firstRow === 'object' && !Array.isArray(firstRow)) {
										extractedHeaders = Object.keys(firstRow);
									} else if (Array.isArray(firstRow)) {
										extractedHeaders = firstRow;
									}
								}
								const info = getImportType(extractedHeaders);
								setFileInfos((prev2) => {
									const copy2 = Array.isArray(prev2) ? [...prev2] : [];
									copy2[newIndex] = info;
									return copy2;
								});
							} catch (err) {
								// ignore parse errors for the chunk
							}
						};
						r.onerror = () => { /* ignore */ };
						r.readAsText(f.slice(0, 64 * 1024));

						// pick the first CSV among the newly appended files to parse and make active (back-compat)
						if (!parsedNewCsv) {
							parsedNewCsv = true;
							parseCSVFile(f, newIndex);
						}
					}
				}
			});

			return copy;
		});
	};

	const handleFile = (file) => {
		// keep backward-compatible behavior: add single file and parse if CSV
		if (!file) return;
		handleAddFiles([file]);
	};

	// derive columns array to pass to getImportType (names as they appear)
	const columns = headers;

	// compute import info once and memoize based on headers so it stays stable across renders
	const importInfo = useMemo(() => getImportType(columns), [Array.isArray(columns) ? columns.join('|') : String(columns)]);

	// merge matched array with any rename targets so consumers see renamed columns too
	const matchedWithRenames = useMemo(() => {
		const base = Array.isArray(importInfo && importInfo.matched) ? [...importInfo.matched] : [];
		const rename = importInfo && importInfo.rename;
		if (rename && typeof rename === 'object') {
			Object.keys(rename).forEach((orig) => {
				const target = rename[orig];
				if (target && base.indexOf(target) === -1) {
					base.push(target);
				}
			});
		}
		return base;
	}, [importInfo && JSON.stringify(importInfo)]);

	// Add: helper to apply rename mapping to headers and parsed data
	const applyRenames = (dataInput, headersInput, renameMap) => {
		if (!renameMap || typeof renameMap !== 'object') return { data: dataInput, headers: headersInput };
		const normalize = (v) => (v === null || v === undefined ? '' : String(v).toLowerCase().trim());
		const normMap = {};
		Object.keys(renameMap).forEach((k) => { normMap[normalize(k)] = renameMap[k]; });

		const newHeaders = Array.isArray(headersInput)
			? headersInput.map((h) => (normMap[normalize(h)] ? normMap[normalize(h)] : h))
			: headersInput;

		let newData = dataInput;
		if (Array.isArray(dataInput) && dataInput.length > 0) {
			const first = dataInput[0];
			// only rename object rows (array-of-objects); leave array rows as-is (header row handling stays as before)
			if (first && typeof first === 'object' && !Array.isArray(first)) {
				newData = dataInput.map((row) => {
					const out = {};
					Object.keys(row).forEach((k) => {
						const nk = normMap[normalize(k)] || k;
						out[nk] = row[k];
					});
					return out;
				});
			}
		}

		return { data: newData, headers: newHeaders };
	};

	// Memoize renamed data/headers so DataProcessor gets stable references (prevents infinite loop)
	const renamed = useMemo(() => {
		return applyRenames(parsedData, headers, importInfo && importInfo.rename);
		// stringify rename to detect changes; headers/parsedData references are used directly
	}, [parsedData, Array.isArray(headers) ? headers.join('|') : String(headers), importInfo && JSON.stringify(importInfo.rename)]);

	// effective exclude filter: prop overrides detected filter from getImportType
	const effectiveExcludeFilter = useMemo(() => {
		if (excludeFilter && excludeFilter.column) return excludeFilter;
		if (importInfo && importInfo.excludeFilter && importInfo.excludeFilter.column) return importInfo.excludeFilter;
		return null;
	}, [excludeFilter && JSON.stringify(excludeFilter), importInfo && JSON.stringify(importInfo && importInfo.excludeFilter)]);

	// effective hyperLink: explicit prop overrides detected hyperLink from getImportType
	const effectiveHyperLink = useMemo(() => {
		if (hyperLink && hyperLink.column) return hyperLink;
		if (importInfo && importInfo.hyperLink && importInfo.hyperLink.column) return importInfo.hyperLink;
		return null;
	}, [hyperLink && JSON.stringify(hyperLink), importInfo && JSON.stringify(importInfo && importInfo.hyperLink)]);

	// NEW: ensure the hyperlink column (if defined) is included in the matched set
	const matchedWithLink = useMemo(() => {
		const base = Array.isArray(matchedWithRenames) ? [...matchedWithRenames] : [];
		// prefer the explicit effectiveHyperLink (prop or detected) then fallback to importInfo.hyperLink
		const hyper = effectiveHyperLink || (importInfo && importInfo.hyperLink);
		const col = hyper && hyper.column ? String(hyper.column) : null;
		if (col && base.indexOf(col) === -1) base.push(col);
		return base;
	}, [matchedWithRenames, effectiveHyperLink && JSON.stringify(effectiveHyperLink), importInfo && JSON.stringify(importInfo && importInfo.hyperLink)]);

	// helper to build hyperlinks and ensure the target column exists
	const escapeRegExp = (s) => String(s || '').replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
	const applyHyperlink = (renamedObj, hyper) => {
		if (!hyper || !hyper.column) return renamedObj;
		const colName = hyper.column;
		const template = hyper.linkLocation || '';
		const paramsDef = Array.isArray(hyper.parameter) ? hyper.parameter : [];

		const headersIn = Array.isArray(renamedObj.headers) ? [...renamedObj.headers] : renamedObj.headers;
		let dataIn = renamedObj.data;
		if (!Array.isArray(dataIn) || dataIn.length === 0) {
			// still ensure header exists so downstream writers see the column
			if (Array.isArray(headersIn) && headersIn.indexOf(colName) === -1) headersIn.push(colName);
			return { data: dataIn, headers: headersIn };
		}

		const first = dataIn[0];

		// helper to build excel HYPERLINK formula safely
		const escapeForExcel = (s) => String(s || '').replace(/"/g, '""');
		const makeHyperlinkFormula = (url, friendly) => {
			const u = escapeForExcel(url);
			const f = escapeForExcel(friendly == null || friendly === '' ? url : friendly);
			// Excel formula: =HYPERLINK("url","friendly")
			return `=HYPERLINK("${u}","${f}")`;
		};

		// object rows
		if (first && typeof first === 'object' && !Array.isArray(first)) {
			if (Array.isArray(headersIn) && headersIn.indexOf(colName) === -1) headersIn.push(colName);
			const out = dataIn.map((row) => {
				const newRow = { ...(row || {}) };
				// collect param values by matching param names case-insensitively against row keys
				const paramValues = paramsDef.map((p) => {
					const needle = String(p || '').toLowerCase().trim();
					const foundKey = Object.keys(row).find((k) => String(k || '').toLowerCase().trim() === needle);
					return foundKey ? row[foundKey] : (row[p] || '');
				});

				// try token replacement in template using exact param text matches
				let url = template;
				if (typeof url === 'string' && url.length > 0 && paramsDef.length > 0) {
					paramsDef.forEach((p, i) => {
						const token = String(p || '');
						const val = paramValues[i] == null ? '' : String(paramValues[i]);
						url = url.replace(new RegExp(escapeRegExp(token), 'g'), encodeURIComponent(val));
					});
					// if replacement didn't change template meaningfully, fallback to join
					if (url === template) {
						const suffix = paramValues.map((v) => encodeURIComponent(String(v || ''))).join('/');
						url = template.endsWith('/') ? (template + suffix) : (template + (template.includes('?') ? '&' : '/') + suffix);
					}
				} else {
					// no template: just join params onto empty base
					url = paramsDef.length > 0 ? paramsDef.map((_, i) => encodeURIComponent(String(paramValues[i] || ''))).join('/') : '';
				}

				// wrap with Excel HYPERLINK and use friendlyName when provided
				const friendly = hyper.friendlyName || colName || url;
				newRow[colName] = makeHyperlinkFormula(url, friendly);
				return newRow;
			});
			return { data: out, headers: headersIn };
		}

		// array rows (header row may be present in headersIn)
		if (Array.isArray(first)) {
			let idx = Array.isArray(headersIn) ? headersIn.findIndex((h) => String(h || '').toLowerCase().trim() === String(colName || '').toLowerCase().trim()) : -1;
			if (idx === -1 && Array.isArray(headersIn)) {
				headersIn.push(colName);
				idx = headersIn.length - 1;
			}
			const out = dataIn.map((row) => {
				const newRow = Array.isArray(row) ? [...row] : [];
				// compute param values by matching headersIn
				const paramValues = paramsDef.map((p) => {
					const needle = String(p || '').toLowerCase().trim();
					const foundIdx = headersIn.findIndex((h) => String(h || '').toLowerCase().trim() === needle);
					return foundIdx !== -1 && Array.isArray(row) ? row[foundIdx] : '';
				});
				// build url similar to object case
				let url = template;
				if (typeof url === 'string' && url.length > 0 && paramsDef.length > 0) {
					paramsDef.forEach((p, i) => {
						const token = String(p || '');
						const val = paramValues[i] == null ? '' : String(paramValues[i]);
						url = url.replace(new RegExp(escapeRegExp(token), 'g'), encodeURIComponent(val));
					});
					if (url === template) {
						const suffix = paramValues.map((v) => encodeURIComponent(String(v || ''))).join('/');
						url = template.endsWith('/') ? (template + suffix) : (template + '/' + suffix);
					}
				} else {
					url = paramsDef.length > 0 ? paramsDef.map((_, i) => encodeURIComponent(String(paramValues[i] || ''))).join('/') : '';
				}
				// ensure row has correct length and set HYPERLINK formula at idx
				while (newRow.length < idx) newRow.push('');
				const friendly = hyper.friendlyName || colName || url;
				newRow[idx] = makeHyperlinkFormula(url, friendly);
				return newRow;
			});
			return { data: out, headers: headersIn };
		}

		return renamedObj;
	};

	// enriched renamed data/headers with hyperlink column added (when applicable)
	const enriched = useMemo(() => applyHyperlink(renamed, effectiveHyperLink), [renamed && JSON.stringify(renamed.headers), renamed && (Array.isArray(renamed.data) ? renamed.data.length : JSON.stringify(renamed.data)), effectiveHyperLink && JSON.stringify(effectiveHyperLink)]);

	// Log generated hyperlinks whenever enriched or the effectiveHyperLink changes
	useEffect(() => {
		try {
			if (!enriched) {
				console.log('ImportManager: enriched not ready yet.');
				return;
			}

			const colName = effectiveHyperLink && effectiveHyperLink.column;
			if (!colName) {
				console.log('ImportManager: effectiveHyperLink has no column defined.', effectiveHyperLink);
				return;
			}

			// collect values (sample up to 20) from enriched.data, supporting object-rows and array-rows
			const dataArr = Array.isArray(enriched.data) ? enriched.data : [];
			const sample = dataArr.slice(0, 20).map((row) => {
				// object rows
				if (row && typeof row === 'object' && !Array.isArray(row)) {
					return row[colName];
				}
				// array rows: find index by header name (case-insensitive)
				if (Array.isArray(row)) {
					const idx = Array.isArray(enriched.headers)
						? enriched.headers.findIndex((h) => String(h || '').toLowerCase().trim() === String(colName || '').toLowerCase().trim())
						: -1;
					return idx !== -1 ? row[idx] : undefined;
				}
				return undefined;
			}).filter((v) => v !== undefined);

			console.log(`ImportManager: generated ${sample.length} hyperlink(s) (sample up to 20):`, sample, { column: colName });
		} catch (err) {
			console.error('ImportManager: error while logging generated hyperlinks', err);
		}
	}, [enriched && JSON.stringify(enriched.headers), enriched && (Array.isArray(enriched.data) ? enriched.data.length : JSON.stringify(enriched.data)), effectiveHyperLink && JSON.stringify(effectiveHyperLink)]);

	// helper to exclude rows where a specified column contains a given value (case-insensitive)
	const applyExclusion = (dataInput, headersInput, filter) => {
		if (!filter || !filter.column || (filter.value === null || filter.value === undefined || String(filter.value).trim() === '')) {
			return dataInput;
		}

		const colName = String(filter.column).toLowerCase().trim();
		const excludeText = String(filter.value).toLowerCase();

		// if no data or headers, nothing to do
		if (!Array.isArray(dataInput) || dataInput.length === 0) return dataInput;

		const originalLength = dataInput.length;
		let result = dataInput;

		// object rows (array of objects)
		const first = dataInput[0];
		if (first && typeof first === 'object' && !Array.isArray(first)) {
			result = dataInput.filter((row) => {
				// find matching key in row (case-insensitive)
				for (const k of Object.keys(row)) {
					if (String(k || '').toLowerCase().trim() === colName) {
						const val = row[k];
						if (val === null || val === undefined) return true; // keep if nothing to compare
						return String(val).toLowerCase().indexOf(excludeText) === -1;
					}
				}
				// if column not present in this row, keep the row
				return true;
			});
		} else if (Array.isArray(first)) {
			// array rows (first row may be header row or arrays of values)
			// find column index from headersInput if available
			let idx = -1;
			if (Array.isArray(headersInput) && headersInput.length > 0) {
				idx = headersInput.findIndex((h) => String(h || '').toLowerCase().trim() === colName);
			}
			// if headers didn't help, try scanning first row for matching header name
			if (idx === -1 && Array.isArray(dataInput) && dataInput.length > 0) {
				const headerRow = dataInput[0];
				if (Array.isArray(headerRow)) {
					idx = headerRow.findIndex((h) => String(h || '').toLowerCase().trim() === colName);
				}
			}
			// if we still didn't find an index, nothing to filter
			if (idx !== -1) {
				// filter rows where value at idx contains excludeText
				result = dataInput.filter((row) => {
					const v = row && row[idx];
					if (v === null || v === undefined) return true;
					return String(v).toLowerCase().indexOf(excludeText) === -1;
				});
			}
		}

		const excludedCount = Math.max(0, originalLength - (Array.isArray(result) ? result.length : 0));
		if (excludedCount > 0) {
			/* eslint-disable no-console */
			console.log(`ImportManager: excluded ${excludedCount} row(s) where column "${filter.column}" contains "${filter.value}"`);
			/* eslint-enable no-console */
		}

		return result;
	};

	// memoized filtered data: apply exclusion on top of renamed data
	const filteredData = useMemo(() => {
		return applyExclusion(enriched.data, enriched.headers, effectiveExcludeFilter);
		// depend on enriched result and the effective filter
	}, [enriched && JSON.stringify(enriched.data ? (Array.isArray(enriched.data) ? [enriched.data.length] : enriched.data) : enriched), enriched && JSON.stringify(enriched.headers), effectiveExcludeFilter && JSON.stringify(effectiveExcludeFilter)]);

	// triggered when user clicks the Import button
	const handleImport = () => {
		// if no files uploaded
		if (!uploadedFiles || uploadedFiles.length === 0) {
			setStatus('No file/data to import.');
			return;
		}

		// find first CSV file from top to bottom
		const firstCsvIdx = uploadedFiles.findIndex((f) => f && /\.csv$/i.test(f.name));
		if (firstCsvIdx === -1) {
			setStatus('No CSV files found to import.');
			return;
		}

		// start multi-file processing
		setProcessingIndex(firstCsvIdx);
		setLastEmittedIndex(-1);
		setIsImported(true);
		setStatus(`Starting import 1 of ${uploadedFiles.length}...`);

		// parse and activate the first CSV to kick off processing
		parseCSVFile(uploadedFiles[firstCsvIdx], firstCsvIdx);
	};

	// Emit onImport for current file once parsed and validated, then allow DataProcessor to run.
	useEffect(() => {
		if (!isImported) return;
		if (processingIndex === -1) return;
		// only act when parsedData corresponds to the processingIndex
		if (activeIndex !== processingIndex) return;
		// ensure we call onImport exactly once per file
		if (lastEmittedIndex === processingIndex) return;

		// validate workbook settings & identifiers for this file (same logic as before)
		try {
			const wbSettings = getWorkbookSettings(headers);
			setWorkbookColumns(Array.isArray(wbSettings.columns) ? [...wbSettings.columns] : []);
			if (typeof DataProcessor.logStaticColumns === 'function') {
				DataProcessor.logStaticColumns(wbSettings.columns);
			}
			/* eslint-disable no-console */
			console.log('ImportManager: workbook settings at import ->', wbSettings);
			/* eslint-enable no-console */

			// if settings define identifier columns, ensure this file contains at least one
			const normalize = (v) => (v === null || v === undefined ? '' : String(v).replace(/\s/g, '').toLowerCase());
			const identifierCandidates = Array.isArray(wbSettings.columns)
				? wbSettings.columns.filter((c) => c && (c.identifier || c.identifer))
				: [];
			if (identifierCandidates.length > 0) {
				const headerKeys = Array.isArray(headers) ? headers.map((h) => normalize(h)) : [];
				let foundAny = false;
				for (const cand of identifierCandidates) {
					const candKeys = new Set();
					if (cand.name) candKeys.add(normalize(cand.name));
					const alias = cand.alias;
					if (Array.isArray(alias)) alias.forEach((a) => { if (a) candKeys.add(normalize(a)); });
					else if (alias) candKeys.add(normalize(alias));
					if (headerKeys.some((hk) => candKeys.has(hk))) {
						foundAny = true;
						break;
					}
				}
				if (!foundAny) {
					const first = identifierCandidates[0];
					const idDisplay = first && (first.name || (Array.isArray(first.alias) ? first.alias[0] : first.alias)) || 'identifier';
					const msg = `Import aborted: none of the configured identifier columns (first: "${idDisplay}") were found in the import file.`;
					/* eslint-disable no-console */
					console.warn(msg);
					/* eslint-enable no-console */
					setStatus(msg);
					// abort whole multi-import
					setIsImported(false);
					setProcessingIndex(-1);
					return;
				}
			}
		} catch (err) {
			// ignore settings read errors
		}

		// ready to emit for this file
		const currentFile = uploadedFiles[processingIndex];
		const dataToEmit = filteredData;
		if (typeof onImport === 'function') {
			onImport({
				file: currentFile,
				data: dataToEmit,
				type: importInfo.type || 'csv',
				matched: matchedWithLink,
				headers: enriched.headers || renamed.headers,
			});
		} else {
			console.log('Imported CSV data', {
				file: currentFile,
				data: dataToEmit,
				type: importInfo.type,
				matched: matchedWithLink,
				headers: enriched.headers || renamed.headers,
			});
		}

		// mark emitted so we don't re-emit
		setLastEmittedIndex(processingIndex);
		// signal processing has started for this file (DataProcessor will render)
		setStatus(`Processing import ${processingIndex + 1} of ${uploadedFiles.length} into workbook...`);
	}, [isImported, processingIndex, activeIndex, parsedData, headers, filteredData, importInfo, matchedWithLink, enriched, renamed, uploadedFiles, lastEmittedIndex, onImport]);

	// handler invoked when DataProcessor finishes (success or failure)
	const handleProcessorComplete = (result) => {
		// show status for this file
		if (result && result.success) {
			setStatus(`Import ${processingIndex + 1} completed.`);
		} else {
			const errMsg = result && result.error ? `: ${result.error}` : (result && result.reason ? `: ${result.reason}` : '');
			setStatus(`Import ${processingIndex + 1} failed${errMsg}`);
		}

		// advance to next CSV (top-to-bottom). find next index > processingIndex
		const total = uploadedFiles ? uploadedFiles.length : 0;
		let next = (processingIndex === -1) ? 0 : processingIndex + 1;
		while (next < total && !(/\.csv$/i.test(uploadedFiles[next] && uploadedFiles[next].name))) {
			next += 1;
		}

		if (next < total) {
			// continue with next CSV
			setProcessingIndex(next);
			setLastEmittedIndex(-1);
			// parse next file which will trigger the effect to emit and run DataProcessor
			parseCSVFile(uploadedFiles[next], next);
		} else {
			// finished all files
			setIsImported(false);
			setProcessingIndex(-1);
			setActiveIndex(-1);
			/* eslint-disable no-console */
			console.log('ImportManager: all files processed.');
			/* eslint-enable no-console */
			setStatus('All imports completed.');
		}
	};

	// handler to receive incremental status updates from DataProcessor
	const handleProcessorStatus = (message) => {
		if (typeof message === 'string' && message) setStatus(message);
	};

	// replaced click handler: simplified, open file picker immediately
	const onButtonClick = () => {
		if (inputRef.current) {
			inputRef.current.value = null;
			inputRef.current.click();
		}
	};

	// drag handlers
	const handleDragOver = (e) => {
		e.preventDefault();
		e.stopPropagation();
		if (!dragActive) setDragActive(true);
	};

	const handleDragLeave = (e) => {
		e.preventDefault();
		e.stopPropagation();
		// only clear when leaving the drop-zone element
		setDragActive(false);
	};

	const handleDrop = (e) => {
		e.preventDefault();
		e.stopPropagation();
		setDragActive(false);
		const list = e.dataTransfer && e.dataTransfer.files;
		if (list && list.length > 0) {
			handleAddFiles(list);
		}
	};

	// allow clicking the drop area to open picker
	const openFilePicker = () => {
		if (inputRef.current) {
			inputRef.current.value = null;
			inputRef.current.click();
		}
	};

	return (
		<div style={styles.container}>
			{/* Title at the top */}
			<div style={{ display: 'flex', justifyContent: 'flex-start', alignItems: 'center', marginTop: 8, marginBottom: 8, paddingLeft: 12 }}>
				<h2 style={{ margin: 0, fontSize: 18, fontWeight: 600, color: '#24303f' }}>Select a file to import</h2>
			</div>

			{/* drop-zone: border around icon + button, clickable and supports drag/drop */}
			<div
				role="button"
				tabIndex={0}
				onClick={openFilePicker}
				onKeyDown={(e) => { if (e.key === 'Enter' || e.key === ' ') openFilePicker(); }}
				onDragOver={handleDragOver}
				onDragEnter={handleDragOver}
				onDragLeave={handleDragLeave}
				onDrop={handleDrop}
				style={{
					margin: '0 0 12px',
					width: '100%',
					boxSizing: 'border-box',
					padding: 16,
					borderRadius: 10,
					display: 'flex',
					flexDirection: 'column',
					alignItems: 'center',
					gap: 12,
					cursor: 'pointer',
					border: dragActive ? '2px dashed #2b6cb0' : '2px dashed rgba(43,108,176,0.25)',
					background: dragActive ? 'rgba(43,108,176,0.03)' : 'transparent',
					transition: 'border-color 120ms ease, background 120ms ease',
				}}
				/* Keep the title static so selected filename is not revealed in the UI */
				title={'Click or drop a CSV file here'}
			>
				{/* icon: simplified to use lucide-react Upload icon */}
				<Upload
					size={48}
					color="#92cbf7"
					// keep layout consistent with previous image
					style={{ display: 'block' }}
				/>

				{/* helper text above the choose button */}
				<div style={{ fontSize: 12, color: '#444', marginTop: 4 }}>
					Drag files here or
				</div>

				{/* upload button (shows filename when selected) */}
				<button
					type="button"
					onClick={(e) => { e.stopPropagation(); onButtonClick(); }}
					// preserve shrink animation behavior and ellipsis for long names
					style={{
						background: '#2b6cb0',
						color: '#fff',
						border: 'none',
						cursor: 'pointer',
						padding: '6px 20px',
						borderRadius: 999,
						fontSize: 14,
						fontWeight: 600,
						boxShadow: '0 2px 6px rgba(43,108,176,0.18)',
						width: '100%',
						maxWidth: '100%',
						overflow: 'hidden',
						textOverflow: 'ellipsis',
						whiteSpace: 'nowrap',
						display: 'inline-block',
					}}
					/* Always show the same label and accessible text — do not surface fileName */
					aria-label={'Choose file'}
					title={'Choose CSV file'}
				>
					{'Choose File'}
				</button>
			</div>

			<input
				ref={inputRef}
				type="file"
				accept=".csv"
				multiple
				style={{ display: 'none' }}
				onChange={(e) => {
					const files = e.target.files;
					if (files && files.length > 0) handleAddFiles(files);
				}}
			/>
			{/* uploaded files list */}
			{uploadedFiles && uploadedFiles.length > 0 && (
				<div style={{ width: '100%', boxSizing: 'border-box', margin: '8px 0', padding: '8px 12px' }}>
					<div style={{ marginBottom: 8, fontSize: 13, color: '#24303f', fontWeight: 600 }}>
						Uploaded files
					</div>
					{uploadedFiles.map((f, idx) => (
						<FileCard
							key={`${f.name}-${idx}`}
							file={f}
							rows={idx === activeIndex && Array.isArray(parsedData) ? parsedData.length : undefined}
							type={fileInfos[idx] && fileInfos[idx].type}
							action={fileInfos[idx] && fileInfos[idx].action}
							icon={fileInfos[idx] && fileInfos[idx].icon} // <-- new: pass icon through
						/>
					))}
				</div>
			)}
			{status && (
				<div style={{ ...styles.infoBox }}>
					<div style={styles.statusText}>Status: {status}</div>
				</div>
			)}

			{/* show Import button after parsing but before actual import */}
			{parsedData && !isImported && (
				<div style={styles.controlRow}>
					<button
						type="button"
						onClick={handleImport}
						// ensure the import button fills the taskpane width
						style={{ ...styles.importButton, width: '100%', display: 'inline-flex', alignItems: 'center', justifyContent: 'center' }}
					>
						<span style={{ display: 'inline-flex', alignItems: 'center', gap: 8 }}>
							{'Import Data'}
							<img src={ImportIcon} alt="" style={{ width: 19, height: 19, objectFit: 'contain', marginLeft: 4 }} />
						</span>
					</button>

					{/* preview / processing UI can remain hidden until import */}
				</div>
			)}

			{/* DataProcessor receives data only after user clicked Import (active file) */}
			{isImported && parsedData && activeIndex !== -1 && activeIndex === processingIndex && (
				<div style={styles.processorWrap}>
					{/* pass memoized renamed data/headers into DataProcessor to avoid reruns */}
					<DataProcessor
						data={filteredData}
						sheetName="test"
						headers={enriched.headers || renamed.headers}
						settingsColumns={workbookColumns}
						matched={matchedWithLink}
						action={importInfo.action}
						onComplete={handleProcessorComplete} // <-- existing prop
						onStatus={handleProcessorStatus}    // <-- newly added prop
					/>
				</div>
			)}
		</div>
	);
}