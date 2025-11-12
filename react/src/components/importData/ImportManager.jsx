import React, { useState, useRef } from 'react';
import parseCSV from './Parsers/csv';
import DataProcessor from './DataProcessor';
import styles from './importManagerStyles';
import { getImportType } from './ImportType'; // new import
import { getWorkbookSettings } from '../utility/getSettings'; // <-- added import

export default function ImportManager({ onImport } = {}) {
	const [fileName, setFileName] = useState('');
	const [status, setStatus] = useState('');
	const [parsedData, setParsedData] = useState(null);
	const [headers, setHeaders] = useState([]); // store column headers
	const [uploadedFile, setUploadedFile] = useState(null); // new: store selected File
	const [isImported, setIsImported] = useState(false); // new: whether user clicked Import
	const [workbookColumns, setWorkbookColumns] = useState([]); // new: columns from workbook settings
	const inputRef = useRef(null);

	const handleFile = (file) => {
		if (!file) return;
		setFileName(file.name);
		setStatus('Reading file...');
		setUploadedFile(file);
		setIsImported(false); // reset import flag on new upload

		const isCSV = /\.csv$/i.test(file.name);

		const reader = new FileReader();

		reader.onerror = () => {
			setStatus('Failed to read file.');
		};

		if (isCSV) {
			reader.onload = (e) => {
				try {
					const text = e.target.result;
					const data = parseCSV(text);

					// extract headers robustly for later use
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

					// call workbook settings util with detected columns (will log columns)
					try {
						const wbSettings = getWorkbookSettings(extractedHeaders);
						// store resolved columns from settings into state so we can pass them to DataProcessor
						setWorkbookColumns(Array.isArray(wbSettings.columns) ? [...wbSettings.columns] : []);
						// log static columns if helper exists
						if (typeof DataProcessor.logStaticColumns === 'function') {
							DataProcessor.logStaticColumns(wbSettings.columns);
						}
						/* eslint-disable no-console */
						console.log('ImportManager: workbook settings ->', wbSettings);
						/* eslint-enable no-console */
					} catch (err) {
						// ignore errors from settings read
					}

					setStatus(`Parsed ${Array.isArray(data) ? data.length : 0} rows`);
					setParsedData(data);

					// NOTE: do not call onImport here — wait for explicit "Import" click
				} catch (err) {
					setStatus('Error parsing CSV.');
					console.error(err);
				}
			};
			reader.readAsText(file);
			return;
		}

		setStatus('Unsupported file type. Please select a .csv file.');
	};

	// derive columns array to pass to getImportType (names as they appear)
	const columns = headers;

	// compute import info once per render
	const importInfo = getImportType(columns);

	// triggered when user clicks the Import button
	const handleImport = () => {
		if (!uploadedFile || !parsedData) {
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
				file: uploadedFile,
				data: parsedData,
				// pass detected import type and matched columns
				type: importInfo.type || 'csv',
				matched: importInfo.matched || [],
				headers,
			});
		} else {
			console.log('Imported CSV data', {
				file: uploadedFile,
				data: parsedData,
				type: importInfo.type,
				matched: importInfo.matched,
				headers,
			});
		}
		setStatus('Import completed.');
		setIsImported(true); // allow DataProcessor to receive/process data now
	};

	const onButtonClick = () => {
		if (inputRef.current) inputRef.current.value = null;
		inputRef.current && inputRef.current.click();
	};

	return (
		<div style={styles.container}>
			{/* renamed button to 'Upload' */}
			<button type="button" onClick={onButtonClick} style={styles.uploadButton}>
				Upload
			</button>
			<input
				ref={inputRef}
				type="file"
				accept=".csv"
				style={{ display: 'none' }}
				onChange={(e) => {
					const file = e.target.files && e.target.files[0];
					handleFile(file);
				}}
			/>
			{fileName && (
				<div style={{ ...styles.infoBox }}>
					<div style={styles.fileName}>File: {fileName}</div>
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
					<button type="button" onClick={handleImport} style={styles.importButton}>
						{importInfo.type}
					</button>

					{/* display the action type under the import button */}
					<div
						style={
							styles.actionText
								? styles.actionText
								: { marginTop: 8, fontSize: 12, color: '#444' }
						}
					>
						Action: {importInfo.action || 'Unknown'}
					</div>

					{/* preview / processing UI can remain hidden until import */}
				</div>
			)}

			{/* DataProcessor receives data only after user clicked Import */}
			{isImported && parsedData && (
				<div style={styles.processorWrap}>
					<DataProcessor
						data={parsedData}
						sheetName="test"
						headers={headers}
						settingsColumns={workbookColumns}
						matched={importInfo.matched}
						action={importInfo.action}
					/>
				</div>
			)}
		</div>
	);
}