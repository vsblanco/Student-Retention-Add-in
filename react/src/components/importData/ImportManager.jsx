import React, { useState, useRef } from 'react';
import parseCSV from './Parsers/csv';
import DataProcessor from './DataProcessor';
import styles from './importManagerStyles';
import { getImportType } from './ImportType';
import { getWorkbookSettings } from '../utility/getSettings';
import { Upload } from 'lucide-react';
import FileCard from './FileCard';

export default function ImportManager({ onImport } = {}) {
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

	const inputRef = useRef(null);

	// drag state for drop-zone
	const [dragActive, setDragActive] = useState(false);

	// parse a single CSV file into parsedData/headers and set active
	const parseCSVFile = (file, index) => {
		if (!file) return;
		setFileName(file.name);
		setStatus('Reading file...');
		setIsImported(false);

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
		// append files and expand fileInfos placeholders, then read a small slice of each CSV
		setUploadedFiles((prev) => {
			const start = prev.length;
			const combined = [...prev, ...filesArr];

			// expand fileInfos to keep indexes aligned
			setFileInfos((prevInfos) => [...prevInfos, ...filesArr.map(() => null)]);

			// for each new file, try to read a small slice to detect headers and compute importInfo
			filesArr.forEach((f, i) => {
				const globalIndex = start + i;
				if (/\.csv$/i.test(f.name)) {
					const r = new FileReader();
					// read a chunk (first 64KB) to avoid loading huge files just to detect headers
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
								const copy = [...prev];
								copy[globalIndex] = info;
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

			// choose first CSV among the newly added files to parse and make active (back-compat)
			const idxInNew = filesArr.findIndex((f) => /\.csv$/i.test(f.name));
			if (idxInNew !== -1) {
				const globalIndex = start + idxInNew;
				parseCSVFile(filesArr[idxInNew], globalIndex);
			}

			return combined;
		});
	};

	const handleFile = (file) => {
		// keep backward-compatible behavior: add single file and parse if CSV
		if (!file) return;
		handleAddFiles([file]);
	};

	// derive columns array to pass to getImportType (names as they appear)
	const columns = headers;

	// compute import info once per render
	const importInfo = getImportType(columns);

	// triggered when user clicks the Import button
	const handleImport = () => {
		const activeFile = uploadedFiles[activeIndex];
		if (!activeFile || !parsedData) {
			setStatus('No file/data to import.');
			return;
		}

		// ensure we read workbook settings just before import as well
		try {
			const wbSettings = getWorkbookSettings(headers);
			setWorkbookColumns(Array.isArray(wbSettings.columns) ? [...wbSettings.columns] : []);
			// log static columns at import time as well
			if (typeof DataProcessor.logStaticColumns === 'function') {
				DataProcessor.logStaticColumns(wbSettings.columns);
			}
			/* eslint-disable no-console */
			console.log('ImportManager: workbook settings at import ->', wbSettings);
			/* eslint-enable no-console */

			// NEW: if settings define an identifier, ensure the import file contains that identifier column
			const normalize = (v) => (v === null || v === undefined ? '' : String(v).replace(/\s/g, '').toLowerCase());
			// collect identifier candidates in order and pick the first that exists in the CSV headers
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
					// if any candidate key exists in the file headers, accept this candidate
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
					return; // abort import
				}
			}
		} catch (err) {
			// ignore errors from settings read
		}

		if (typeof onImport === 'function') {
			onImport({
				file: uploadedFiles[activeIndex],
				data: parsedData,
				// pass detected import type and matched columns
				type: importInfo.type || 'csv',
				matched: importInfo.matched || [],
				headers,
			});
		} else {
			console.log('Imported CSV data', {
				file: uploadedFiles[activeIndex],
				data: parsedData,
				type: importInfo.type,
				matched: importInfo.matched,
				headers,
			});
		}

		// show that processing has started; final "completed" status will be set by DataProcessor
		setStatus('Processing import into workbook...');
		setIsImported(true); // allow DataProcessor to receive/process data now

		// NOTE: do NOT set status to 'Import completed.' here — wait for DataProcessor callback
	};

	// handler invoked when DataProcessor finishes (success or failure)
	const handleProcessorComplete = (result) => {
		if (result && result.success) {
			setStatus('Import completed.');
		} else {
			const errMsg = result && result.error ? `: ${result.error}` : (result && result.reason ? `: ${result.reason}` : '');
			setStatus(`Import failed${errMsg}`);
		}
		/* keep DataProcessor visible if you want to inspect the sheet after import */
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
					color="#b5b5b5"
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
						style={{ ...styles.importButton, width: '100%' }}
					>
						{importInfo.type}
					</button>

					{/* preview / processing UI can remain hidden until import */}
				</div>
			)}

			{/* DataProcessor receives data only after user clicked Import (active file) */}
			{isImported && parsedData && activeIndex !== -1 && (
				<div style={styles.processorWrap}>
					<DataProcessor
						data={parsedData}
						sheetName="test"
						headers={headers}
						settingsColumns={workbookColumns}
						matched={importInfo.matched}
						action={importInfo.action}
						onComplete={handleProcessorComplete} // <-- existing prop
						onStatus={handleProcessorStatus}    // <-- newly added prop
					/>
				</div>
			)}
		</div>
	);
}