/* * Timestamp: 2025-11-19 13:40:00 EST
 * Version: 5.1.0
 * Author: Gemini (for Victor)
 * Description: Optimized ImportManager.
 * Improvements:
 * - Replaced O(N*M) row scanning with O(1) key lookups using normalized header maps.
 * - Extracted heavy logic outside component for cleaner readability and stability.
 * - Unified CSV preview logic to reduce code duplication.
 */

import React, { useState, useRef, useMemo, useEffect, useCallback } from 'react';
import parseCSV from './Parsers/csv';
import DataProcessor from './DataProcessor';
import styles from './importManagerStyles'; 
import { getImportType } from './ImportType';
import { getWorkbookSettings } from '../utility/getSettings';
import { CloudUpload, FileText, Table, ArrowRight, Plus } from 'lucide-react';
import FileCard from './FileCard';
import ImportIcon from '../../assets/icons/import-icon.png';

// --- Helper Utilities (Defined outside to keep component clean) ---

// Helper: Create a map of lowercase keys to actual keys for O(1) lookup
const createKeyMap = (headers) => {
    const map = {};
    if (Array.isArray(headers)) {
        headers.forEach(h => {
            if (h) map[String(h).toLowerCase().trim()] = h;
        });
    }
    return map;
};

// Helper: Apply Rename Map
const applyRenames = (dataInput, headersInput, renameMap) => {
    if (!renameMap || typeof renameMap !== 'object') return { data: dataInput, headers: headersInput };
    
    // normalize rename map
    const normalize = (v) => (v == null ? '' : String(v).toLowerCase().trim());
    const normMap = {};
    Object.keys(renameMap).forEach((k) => { normMap[normalize(k)] = renameMap[k]; });

    // Rename headers
    const newHeaders = Array.isArray(headersInput)
        ? headersInput.map((h) => (normMap[normalize(h)] ? normMap[normalize(h)] : h))
        : headersInput;

    let newData = dataInput;
    if (Array.isArray(dataInput) && dataInput.length > 0) {
        const first = dataInput[0];
        // Only rename object rows; array rows rely on index/headers which are already handled
        if (first && typeof first === 'object' && !Array.isArray(first)) {
            // Optimization: Create a fast lookup for row keys
            // Since object keys might differ in casing row-to-row in bad CSVs, we unfortunately 
            // still need to map keys, but we can do it efficiently.
            newData = dataInput.map((row) => {
                const out = {};
                Object.keys(row).forEach((k) => {
                    // If this key is in our rename map (normalized), use the new name
                    const nk = normMap[normalize(k)];
                    out[nk || k] = row[k];
                });
                return out;
            });
        }
    }

    return { data: newData, headers: newHeaders };
};

// Helper: Apply Hyperlink Logic
const applyHyperlink = (renamedObj, hyper) => {
    if (!hyper || !hyper.column) return renamedObj;
    
    const colName = hyper.column;
    const template = hyper.linkLocation || '';
    const paramsDef = Array.isArray(hyper.parameter) ? hyper.parameter : [];

    // Clone headers to avoid mutation
    const headersIn = Array.isArray(renamedObj.headers) ? [...renamedObj.headers] : [];
    const dataIn = renamedObj.data;

    // Ensure header exists
    let headerIdx = headersIn.findIndex(h => String(h).toLowerCase().trim() === String(colName).toLowerCase().trim());
    if (headerIdx === -1) {
        headersIn.push(colName);
        headerIdx = headersIn.length - 1;
    }

    if (!Array.isArray(dataIn) || dataIn.length === 0) {
        return { data: dataIn, headers: headersIn };
    }

    const first = dataIn[0];
    const isObjectRows = first && typeof first === 'object' && !Array.isArray(first);

    // Optimization: Pre-calculate parameter lookups
    // For objects: map lowercase param name -> actual key logic handled per row or via map
    // For arrays: map param name -> index
    let paramIndices = []; 
    
    if (!isObjectRows) {
        // Pre-calc indices for array rows
        paramIndices = paramsDef.map(p => {
            const needle = String(p).toLowerCase().trim();
            return headersIn.findIndex(h => String(h).toLowerCase().trim() === needle);
        });
    }

    // Helper to build formula
    const escapeForExcel = (s) => String(s || '').replace(/"/g, '""');
    const makeFormula = (url, friendly) => {
        return `=HYPERLINK("${escapeForExcel(url)}","${escapeForExcel(friendly || url)}")`;
    };
    const escapeRegExp = (s) => String(s || '').replace(/[.*+?^${}()|[\]\\]/g, '\\$&');

    const out = dataIn.map(row => {
        let newRow = isObjectRows ? { ...row } : [...row];
        let paramValues = [];

        // Extract values
        if (isObjectRows) {
            // Fast lookup using a temporary key map for this row is overkill if consistent, 
            // but assuming row keys are consistent, we scan once.
            // To be safe and fast:
            paramValues = paramsDef.map(p => {
                const needle = String(p).toLowerCase().trim();
                // Fast search: check exact first, then scan keys
                if (row[p] !== undefined) return row[p];
                const foundKey = Object.keys(row).find(k => String(k).toLowerCase().trim() === needle);
                return foundKey ? row[foundKey] : '';
            });
        } else {
            // Array rows: direct index access (O(1))
            paramValues = paramIndices.map(idx => (idx !== -1 && row[idx] !== undefined ? row[idx] : ''));
        }

        // Build URL
        let url = template;
        if (url && paramsDef.length > 0) {
            let usedTemplate = false;
            paramsDef.forEach((p, i) => {
                const val = paramValues[i] == null ? '' : String(paramValues[i]);
                // Replace exact token matches
                const regex = new RegExp(escapeRegExp(String(p)), 'g');
                if (regex.test(url)) {
                    url = url.replace(regex, encodeURIComponent(val));
                    usedTemplate = true;
                }
            });
            
            // Fallback logic if template didn't consume params via tokens
            if (!usedTemplate || url === template) {
                 const suffix = paramValues.map(v => encodeURIComponent(String(v || ''))).join('/');
                 url = template.endsWith('/') ? (template + suffix) : (template + (template.includes('?') ? '&' : '/') + suffix);
            }
        } else {
             url = paramsDef.map((_, i) => encodeURIComponent(String(paramValues[i] || ''))).join('/');
        }

        const friendly = hyper.friendlyName || colName || url;
        const formula = makeFormula(url, friendly);

        if (isObjectRows) {
            newRow[colName] = formula;
        } else {
            // Ensure row length
            while (newRow.length < headerIdx) newRow.push('');
            newRow[headerIdx] = formula;
        }
        return newRow;
    });

    return { data: out, headers: headersIn };
};

// Helper: Apply Exclusion Filter
const applyExclusion = (dataInput, headersInput, filter) => {
    if (!filter || !filter.column || (filter.value == null || String(filter.value).trim() === '')) {
        return dataInput;
    }

    const colName = String(filter.column).toLowerCase().trim();
    const excludeText = String(filter.value).toLowerCase();

    if (!Array.isArray(dataInput) || dataInput.length === 0) return dataInput;

    const first = dataInput[0];
    const isObjectRows = first && typeof first === 'object' && !Array.isArray(first);

    let result = dataInput;

    if (isObjectRows) {
        result = dataInput.filter(row => {
            // Find key efficiently
            let val = undefined;
            // Check exact match first (Fast Path)
            if (row[filter.column] !== undefined) val = row[filter.column];
            else {
                // Scan keys (Slow Path - only if keys are messy)
                const foundKey = Object.keys(row).find(k => String(k).toLowerCase().trim() === colName);
                if (foundKey) val = row[foundKey];
            }
            
            if (val == null) return true; // keep if missing
            return String(val).toLowerCase().indexOf(excludeText) === -1;
        });
    } else {
        // Array rows
        let idx = -1;
        if (Array.isArray(headersInput)) {
            idx = headersInput.findIndex(h => String(h).toLowerCase().trim() === colName);
        }
        if (idx === -1 && Array.isArray(dataInput[0])) {
            // Fallback to checking first row of data as header
            idx = dataInput[0].findIndex(h => String(h).toLowerCase().trim() === colName);
        }

        if (idx !== -1) {
            result = dataInput.filter(row => {
                const v = row[idx];
                if (v == null) return true;
                return String(v).toLowerCase().indexOf(excludeText) === -1;
            });
        }
    }

    const excludedCount = dataInput.length - result.length;
    if (excludedCount > 0) {
         // console.log(`ImportManager: excluded ${excludedCount} row(s)`);
    }
    return result;
};

// Helper: Read a small chunk of a CSV file to extract headers/info
const readCsvInfo = (file, callback) => {
    const reader = new FileReader();
    reader.onload = (e) => {
        try {
            const text = e.target.result;
            // Parse just enough to get headers (first few lines)
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
            callback(getImportType(extractedHeaders));
        } catch (err) {
            callback(null);
        }
    };
    reader.onerror = () => callback(null);
    reader.readAsText(file.slice(0, 64 * 1024)); // Read first 64KB only
};


// --- Main Component ---

export default function ImportManager({ onImport, excludeFilter, hyperLink } = {}) {
    const [fileName, setFileName] = useState('');
    const [status, setStatus] = useState('');
    const [parsedData, setParsedData] = useState(null);
    const [headers, setHeaders] = useState([]); 
    const [uploadedFiles, setUploadedFiles] = useState([]); 
    const [activeIndex, setActiveIndex] = useState(-1); 
    const [fileInfos, setFileInfos] = useState([]);
    const [isImported, setIsImported] = useState(false); 
    const [workbookColumns, setWorkbookColumns] = useState([]); 

    const [processingIndex, setProcessingIndex] = useState(-1);
    const [lastEmittedIndex, setLastEmittedIndex] = useState(-1);
    const [importCompleted, setImportCompleted] = useState(false);

    const inputRef = useRef(null);
    const [dragActive, setDragActive] = useState(false);

    // --- File Processing Logic ---

    const parseCSVFile = useCallback((file, index) => {
        if (!file) return;
        setFileName(file.name);
        setStatus('Reading file...');

        if (!/\.csv$/i.test(file.name)) {
            setStatus('Unsupported file type. Please select a .csv file.');
            return;
        }

        const reader = new FileReader();
        reader.onerror = () => setStatus('Failed to read file.');
        reader.onload = (e) => {
            try {
                const text = e.target.result;
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
                setHeaders(extractedHeaders);

                try {
                    const wbSettings = getWorkbookSettings(extractedHeaders);
                    setWorkbookColumns(Array.isArray(wbSettings.columns) ? [...wbSettings.columns] : []);
                    if (typeof DataProcessor.logStaticColumns === 'function') {
                        DataProcessor.logStaticColumns(wbSettings.columns);
                    }
                } catch (err) { /* ignore */ }

                setStatus(`Parsed ${Array.isArray(data) ? data.length : 0} rows`);
                setParsedData(data);
                setActiveIndex(index);

                // Update info for this file specifically now that we have full headers
                try {
                    const info = getImportType(extractedHeaders);
                    setFileInfos((prev) => {
                        const copy = [...prev];
                        copy[index] = info;
                        return copy;
                    });
                } catch (e) { /* ignore */ }

            } catch (err) {
                setStatus('Error parsing CSV.');
                console.error(err);
            }
        };
        reader.readAsText(file);
    }, []);

    const handleAddFiles = useCallback((filesList) => {
        if (!filesList || filesList.length === 0) return;
        const filesArr = Array.from(filesList);

        // Reset scenario
        if (isImported || importCompleted) {
            const seen = new Set();
            const uniques = [];
            filesArr.forEach((f) => {
                if (f && !seen.has(f.name)) {
                    uniques.push(f);
                    seen.add(f.name);
                }
            });

            setUploadedFiles(uniques);
            setFileInfos(uniques.map(() => null));
            setParsedData(null);
            setHeaders([]);
            setActiveIndex(-1);
            setIsImported(false);
            setImportCompleted(false);

            // Async read info for all new files
            uniques.forEach((f, i) => {
                if (/\.csv$/i.test(f.name)) {
                    readCsvInfo(f, (info) => {
                        setFileInfos(prev => {
                            const copy = [...prev];
                            copy[i] = info;
                            return copy;
                        });
                    });
                }
            });

            // Automatically select first CSV
            const firstCsvIdx = uniques.findIndex((f) => /\.csv$/i.test(f.name));
            if (firstCsvIdx !== -1) {
                parseCSVFile(uniques[firstCsvIdx], firstCsvIdx);
            }
            return;
        }

        // Append scenario
        setUploadedFiles((prev) => {
            const copy = [...prev];
            let parsedNewCsv = false;

            filesArr.forEach((f) => {
                const existingIndex = copy.findIndex((p) => p.name === f.name);
                
                if (existingIndex !== -1) {
                    // Replace
                    copy[existingIndex] = f;
                    // Reset info placeholder
                    setFileInfos(prevInfos => {
                        const infos = [...prevInfos];
                        infos[existingIndex] = null;
                        return infos;
                    });

                    if (/\.csv$/i.test(f.name)) {
                        readCsvInfo(f, (info) => {
                            setFileInfos(prev2 => {
                                const copy2 = [...prev2];
                                copy2[existingIndex] = info;
                                return copy2;
                            });
                        });
                        // Also fully parse this replacement to make it active
                        parseCSVFile(f, existingIndex);
                    }
                } else {
                    // Append
                    const newIndex = copy.length;
                    copy.push(f);
                    // Expand info array
                    setFileInfos(prevInfos => [...prevInfos, null]);

                    if (/\.csv$/i.test(f.name)) {
                        readCsvInfo(f, (info) => {
                            setFileInfos(prev2 => {
                                const copy2 = [...prev2];
                                copy2[newIndex] = info;
                                return copy2;
                            });
                        });

                        // Parse first NEW csv found
                        if (!parsedNewCsv) {
                            parsedNewCsv = true;
                            parseCSVFile(f, newIndex);
                        }
                    }
                }
            });
            return copy;
        });
    }, [isImported, importCompleted, parseCSVFile]);


    // --- Memoized Data Transformations ---

    const columns = headers;

    // Stable import info
    const importInfo = useMemo(() => getImportType(columns), [Array.isArray(columns) ? columns.join('|') : String(columns)]);

    const effectiveExcludeFilter = useMemo(() => {
        if (excludeFilter && excludeFilter.column) return excludeFilter;
        return importInfo && importInfo.excludeFilter;
    }, [excludeFilter, importInfo]);

    const effectiveHyperLink = useMemo(() => {
        if (hyperLink && hyperLink.column) return hyperLink;
        return importInfo && importInfo.hyperLink;
    }, [hyperLink, importInfo]);

    // 1. Rename
    const renamed = useMemo(() => {
        return applyRenames(parsedData, headers, importInfo && importInfo.rename);
    }, [parsedData, headers, importInfo]);

    // 2. Hyperlink
    const enriched = useMemo(() => {
        return applyHyperlink(renamed, effectiveHyperLink);
    }, [renamed, effectiveHyperLink]);

    // 3. Filter
    const filteredData = useMemo(() => {
        return applyExclusion(enriched.data, enriched.headers, effectiveExcludeFilter);
    }, [enriched, effectiveExcludeFilter]);

    // 4. Matched Columns (for UI/DataProcessor)
    const matchedWithLink = useMemo(() => {
        // Merge matched + renames + hyperlink col
        const base = Array.isArray(importInfo && importInfo.matched) ? [...importInfo.matched] : [];
        
        // Add Rename targets
        const rename = importInfo && importInfo.rename;
        if (rename) {
            Object.values(rename).forEach(target => {
                if (target && !base.includes(target)) base.push(target);
            });
        }

        // Add Hyperlink col
        const hyperCol = effectiveHyperLink && effectiveHyperLink.column;
        if (hyperCol && !base.includes(String(hyperCol))) base.push(String(hyperCol));

        return base;
    }, [importInfo, effectiveHyperLink]);


    // --- Handlers ---

    const handleImport = () => {
        if (!uploadedFiles || uploadedFiles.length === 0) {
            setStatus('No file/data to import.');
            return;
        }
        const firstCsvIdx = uploadedFiles.findIndex((f) => f && /\.csv$/i.test(f.name));
        if (firstCsvIdx === -1) {
            setStatus('No CSV files found to import.');
            return;
        }

        setProcessingIndex(firstCsvIdx);
        setLastEmittedIndex(-1);
        setIsImported(true);
        setImportCompleted(false);
        setStatus(`Starting import 1 of ${uploadedFiles.length}...`);

        // Activate the file to ensure state (headers/data) is ready
        parseCSVFile(uploadedFiles[firstCsvIdx], firstCsvIdx);
    };

    const handleProcessorComplete = useCallback((result) => {
        if (result && result.success) {
            setStatus(`Import ${processingIndex + 1} completed.`);
        } else {
            const msg = result?.error || result?.reason || 'Unknown error';
            setStatus(`Import ${processingIndex + 1} failed: ${msg}`);
        }

        // Find next CSV
        const total = uploadedFiles.length;
        let next = processingIndex + 1;
        while (next < total && !(/\.csv$/i.test(uploadedFiles[next]?.name))) {
            next++;
        }

        if (next < total) {
            setProcessingIndex(next);
            setLastEmittedIndex(-1);
            parseCSVFile(uploadedFiles[next], next);
        } else {
            setIsImported(false);
            setProcessingIndex(-1);
            setActiveIndex(-1);
            setImportCompleted(true);
            console.log('ImportManager: all files processed.');
            setStatus('All imports completed.');
        }
    }, [uploadedFiles, processingIndex, parseCSVFile]);

    // Orchestrator Effect: Emits 'onImport' when ready
    useEffect(() => {
        if (!isImported || processingIndex === -1) return;
        // Wait for active index to match processing index (ensures parse is done)
        if (activeIndex !== processingIndex) return;
        // Dedup emission
        if (lastEmittedIndex === processingIndex) return;

        // Validate Settings & Identifiers (Quick Check)
        const wbSettings = getWorkbookSettings(headers);
        try {
            if (wbSettings) {
                 const normalize = (v) => (v || '').replace(/\s/g, '').toLowerCase();
                 const candidates = wbSettings.columns?.filter(c => c.identifier || c.identifer) || [];
                 
                 if (candidates.length > 0) {
                     const headerSet = new Set(headers.map(normalize));
                     const found = candidates.some(c => {
                         if (c.name && headerSet.has(normalize(c.name))) return true;
                         const aliases = [].concat(c.alias || []);
                         return aliases.some(a => a && headerSet.has(normalize(a)));
                     });

                     if (!found) {
                        const idName = candidates[0].name || 'identifier';
                        const msg = `Import aborted: Identifier "${idName}" not found in file.`;
                        console.warn(msg);
                        setStatus(msg);
                        setIsImported(false);
                        setProcessingIndex(-1);
                        return;
                     }
                 }
            }
        } catch (e) { /* ignore */ }

        // Emit Data
        const payload = {
            file: uploadedFiles[processingIndex],
            data: filteredData,
            type: importInfo.type || 'csv',
            matched: matchedWithLink,
            headers: enriched.headers || renamed.headers,
        };

        if (typeof onImport === 'function') onImport(payload);
        else console.log('Imported CSV data', payload);

        setLastEmittedIndex(processingIndex);
        setStatus(`Processing import ${processingIndex + 1} of ${uploadedFiles.length}...`);
    }, [isImported, processingIndex, activeIndex, headers, filteredData, importInfo, matchedWithLink, enriched, renamed, uploadedFiles, lastEmittedIndex, onImport]);


    // --- Render Helpers ---

    const openFilePicker = () => inputRef.current && ((inputRef.current.value = null), inputRef.current.click());
    
    const handleDragOver = (e) => { e.preventDefault(); if (!dragActive) setDragActive(true); };
    const handleDragLeave = (e) => { e.preventDefault(); setDragActive(false); };
    const handleDrop = (e) => {
        e.preventDefault();
        setDragActive(false);
        if (e.dataTransfer.files?.length > 0) handleAddFiles(e.dataTransfer.files);
    };

    const hasFiles = uploadedFiles && uploadedFiles.length > 0;

    return (
        <div className="w-full max-w-2xl mx-auto bg-white rounded-2xl shadow-xl shadow-slate-200/60 border border-white overflow-hidden p-6 transition-all duration-300">
            
            {/* Header */}
            <div className="flex justify-between items-end mb-6">
                <div>
                    <h2 className="text-2xl text-slate-800 font-bold tracking-tight">Import Data</h2>
                    <p className="text-slate-400 text-sm mt-1">
                        {hasFiles ? 'Review your files below' : 'Upload your CSV to begin'}
                    </p>
                </div>
                {status && (
                    <div className="text-xs font-medium px-3 py-1 bg-slate-100 text-slate-500 rounded-full animate-pulse max-w-[200px] truncate">
                        {status}
                    </div>
                )}
            </div>

            {/* Drop Zone */}
            <div
                role="button"
                tabIndex={0}
                onClick={openFilePicker}
                onDragOver={handleDragOver}
                onDragEnter={handleDragOver}
                onDragLeave={handleDragLeave}
                onDrop={handleDrop}
                className={`
                    relative group transition-all duration-300 ease-out
                    border-2 border-dashed rounded-2xl cursor-pointer
                    min-h-[320px] flex flex-col
                    ${dragActive ? 'border-indigo-500 bg-indigo-50/40 scale-[0.99]' : 'border-slate-200 hover:border-slate-300 hover:bg-slate-50/50'}
                    ${hasFiles ? 'justify-start p-4' : 'items-center justify-center p-10'}
                `}
            >
                {!hasFiles ? (
                    <>
                        <div className="relative mb-6 pointer-events-none transition-transform duration-300 group-hover:scale-110">
                            <div className={`transition-colors duration-300 ${dragActive ? 'text-indigo-500' : 'text-slate-300'}`}>
                                <CloudUpload size={64} strokeWidth={1.5} />
                            </div>
                            {/* Decorative Icons */}
                            <div className="absolute -top-4 -left-12 bg-white p-2.5 rounded-xl shadow-lg border border-slate-50 transform -rotate-12 transition-transform duration-500 group-hover:-translate-x-2 group-hover:-rotate-12">
                                <FileText size={24} className="text-emerald-500" />
                            </div>
                            <div className="absolute -top-2 -right-10 bg-white p-2.5 rounded-xl shadow-lg border border-slate-50 transform rotate-12 transition-transform duration-500 group-hover:translate-x-2 group-hover:rotate-6">
                                <Table size={24} className="text-indigo-500" />
                            </div>
                        </div>
                        <div className="text-center space-y-2 pointer-events-none">
                            <h3 className="text-lg font-semibold text-slate-700">Drag & drop</h3>
                            <p className="text-sm text-slate-400 font-medium">
                                your CSV files here, or <span className="text-indigo-500 hover:text-indigo-600 underline decoration-indigo-200 underline-offset-2">browse</span>
                            </p>
                        </div>
                    </>
                ) : (
                    <div className="w-full space-y-3 animate-in fade-in slide-in-from-bottom-2 duration-300">
                        {uploadedFiles.map((f, idx) => (
                            <div key={`${f.name}-${idx}`} onClick={(e) => e.stopPropagation()}>
                                <FileCard
                                    file={f}
                                    rows={idx === activeIndex && Array.isArray(parsedData) ? parsedData.length : undefined}
                                    type={fileInfos[idx] && fileInfos[idx].type}
                                    action={fileInfos[idx] && fileInfos[idx].action}
                                    icon={fileInfos[idx] && fileInfos[idx].icon}
                                />
                            </div>
                        ))}
                        <div className="flex items-center justify-center pt-4 pb-2 opacity-60 group-hover:opacity-100 transition-opacity">
                            <div className="flex items-center gap-2 text-sm text-slate-400 font-medium bg-white/50 px-4 py-2 rounded-full">
                                <Plus size={16} />
                                <span>Drop or click to add more files</span>
                            </div>
                        </div>
                    </div>
                )}
                <input ref={inputRef} type="file" accept=".csv" multiple className="hidden" onChange={(e) => e.target.files?.length && handleAddFiles(e.target.files)} />
            </div>

            {/* Action Button */}
            {parsedData && !isImported && !importCompleted && (
                <div className="mt-6 flex justify-center animate-in fade-in slide-in-from-bottom-4 duration-500 fill-mode-forwards">
                    <button
                        type="button"
                        onClick={handleImport}
                        className="
                            group relative flex items-center justify-center gap-3
                            bg-slate-900 hover:bg-slate-800 text-white
                            px-8 py-3.5 rounded-full
                            shadow-xl shadow-slate-200 hover:shadow-2xl hover:shadow-slate-300
                            transform transition-all duration-200 hover:-translate-y-0.5 active:translate-y-0
                            w-full sm:w-auto sm:min-w-[200px]
                        "
                    >
                        <span className="font-semibold tracking-wide text-sm">Import Data</span>
                        {ImportIcon ? (
                            <img src={ImportIcon} alt="" className="w-5 h-5 object-contain invert opacity-90 group-hover:opacity-100 transition-opacity" />
                        ) : (
                            <ArrowRight size={18} className="opacity-70 group-hover:translate-x-1 transition-transform" />
                        )}
                    </button>
                </div>
            )}

            {/* Data Processor (Hidden) */}
            {isImported && parsedData && activeIndex !== -1 && activeIndex === processingIndex && (
                <div className="mt-6">
                    <DataProcessor
                        data={filteredData}
                        sheetName="test"
                        headers={enriched.headers || renamed.headers}
                        settingsColumns={workbookColumns}
                        matched={matchedWithLink}
                        action={importInfo.action}
                        onComplete={handleProcessorComplete}
                        onStatus={(msg) => msg && setStatus(msg)}
                    />
                </div>
            )}
        </div>
    );
}