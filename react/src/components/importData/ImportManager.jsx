/* * Timestamp: 2025-11-19 15:15:00 EST
 * Version: 5.6.0
 * Author: Gemini (for Victor)
 * Description: Optimized ImportManager.
 * Improvements:
 * - Switched Excel parser to use 'exceljs' (Async) to resolve security vulnerabilities in 'xlsx'.
 * - Refactored file reading logic to handle asynchronous parsing operations.
 */

import React, { useState, useRef, useMemo, useEffect, useCallback } from 'react';
import parseCSV from './Parsers/csv';
import parseExcel from './Parsers/xlsx'; // Now uses ExcelJS (Async)
import DataProcessor from './DataProcessor';
import styles from './importManagerStyles'; 
import { getImportType, IMPORT_DEFINITIONS } from './ImportType';
import { getWorkbookSettings } from '../utility/getSettings';
import { CloudUpload, FileText, Table, ArrowRight, Plus, Info, ChevronDown, ChevronUp } from 'lucide-react';
import FileCard from './FileCard';
import ImportIcon from '../../assets/icons/import-icon.png';

// --- Helper Utilities ---

const createKeyMap = (headers) => {
    const map = {};
    if (Array.isArray(headers)) {
        headers.forEach(h => {
            if (h) map[String(h).toLowerCase().trim()] = h;
        });
    }
    return map;
};

const applyRenames = (dataInput, headersInput, renameMap) => {
    if (!renameMap || typeof renameMap !== 'object') return { data: dataInput, headers: headersInput };
    
    const normalize = (v) => (v == null ? '' : String(v).toLowerCase().trim());
    const normMap = {};
    Object.keys(renameMap).forEach((k) => { normMap[normalize(k)] = renameMap[k]; });

    const newHeaders = Array.isArray(headersInput)
        ? headersInput.map((h) => (normMap[normalize(h)] ? normMap[normalize(h)] : h))
        : headersInput;

    let newData = dataInput;
    if (Array.isArray(dataInput) && dataInput.length > 0) {
        const first = dataInput[0];
        if (first && typeof first === 'object' && !Array.isArray(first)) {
            newData = dataInput.map((row) => {
                const out = {};
                Object.keys(row).forEach((k) => {
                    const nk = normMap[normalize(k)];
                    out[nk || k] = row[k];
                });
                return out;
            });
        }
    }

    return { data: newData, headers: newHeaders };
};

const applyHyperlink = (renamedObj, hyper) => {
    if (!hyper || !hyper.column) return renamedObj;
    
    const colName = hyper.column;
    const template = hyper.linkLocation || '';
    const paramsDef = Array.isArray(hyper.parameter) ? hyper.parameter : [];

    const headersIn = Array.isArray(renamedObj.headers) ? [...renamedObj.headers] : [];
    const dataIn = renamedObj.data;

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

    let paramIndices = []; 
    if (!isObjectRows) {
        paramIndices = paramsDef.map(p => {
            const needle = String(p).toLowerCase().trim();
            return headersIn.findIndex(h => String(h).toLowerCase().trim() === needle);
        });
    }

    const escapeForExcel = (s) => String(s || '').replace(/"/g, '""');
    const makeFormula = (url, friendly) => {
        return `=HYPERLINK("${escapeForExcel(url)}","${escapeForExcel(friendly || url)}")`;
    };
    const escapeRegExp = (s) => String(s || '').replace(/[.*+?^${}()|[\]\\]/g, '\\$&');

    const out = dataIn.map(row => {
        let newRow = isObjectRows ? { ...row } : [...row];
        let paramValues = [];

        if (isObjectRows) {
            paramValues = paramsDef.map(p => {
                const needle = String(p).toLowerCase().trim();
                if (row[p] !== undefined) return row[p];
                const foundKey = Object.keys(row).find(k => String(k).toLowerCase().trim() === needle);
                return foundKey ? row[foundKey] : '';
            });
        } else {
            paramValues = paramIndices.map(idx => (idx !== -1 && row[idx] !== undefined ? row[idx] : ''));
        }

        let url = template;
        if (url && paramsDef.length > 0) {
            let usedTemplate = false;
            paramsDef.forEach((p, i) => {
                const val = paramValues[i] == null ? '' : String(paramValues[i]);
                const regex = new RegExp(escapeRegExp(String(p)), 'g');
                if (regex.test(url)) {
                    url = url.replace(regex, encodeURIComponent(val));
                    usedTemplate = true;
                }
            });
            
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
            while (newRow.length < headerIdx) newRow.push('');
            newRow[headerIdx] = formula;
        }
        return newRow;
    });

    return { data: out, headers: headersIn };
};

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
            let val = undefined;
            if (row[filter.column] !== undefined) val = row[filter.column];
            else {
                const foundKey = Object.keys(row).find(k => String(k).toLowerCase().trim() === colName);
                if (foundKey) val = row[foundKey];
            }
            
            if (val == null) return true;
            return String(val).toLowerCase().indexOf(excludeText) === -1;
        });
    } else {
        let idx = -1;
        if (Array.isArray(headersInput)) {
            idx = headersInput.findIndex(h => String(h).toLowerCase().trim() === colName);
        }
        if (idx === -1 && Array.isArray(dataInput[0])) {
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
    return result;
};

// Helper: Determine file info (headers, type) for either CSV or Excel
// Updated to handle Async Excel parsing
const detectFileInfo = (file, callback) => {
    const isCsv = /\.csv$/i.test(file.name);
    const isExcel = /\.xlsx?$/i.test(file.name);

    if (!isCsv && !isExcel) {
        callback(null);
        return;
    }

    const reader = new FileReader();
    
    reader.onload = async (e) => {
        try {
            let extractedHeaders = [];
            
            if (isCsv) {
                const text = e.target.result;
                const data = parseCSV(text); 
                if (Array.isArray(data) && data.length > 0) {
                    const firstRow = data[0];
                    extractedHeaders = (firstRow && typeof firstRow === 'object' && !Array.isArray(firstRow)) 
                        ? Object.keys(firstRow) 
                        : (Array.isArray(firstRow) ? firstRow : []);
                }
            } else {
                // Async Excel Parse
                const buffer = e.target.result;
                const data = await parseExcel(buffer);
                if (Array.isArray(data) && data.length > 0) {
                    const firstRow = data[0];
                    extractedHeaders = Object.keys(firstRow);
                }
            }
            
            callback(getImportType(extractedHeaders));
        } catch (err) {
            console.error('Error in detectFileInfo:', err);
            callback(null);
        }
    };
    
    reader.onerror = () => callback(null);

    // Read CSV as text, Excel as ArrayBuffer
    if (isCsv) {
        reader.readAsText(file.slice(0, 64 * 1024)); 
    } else {
        reader.readAsArrayBuffer(file);
    }
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
    
    const [showInfo, setShowInfo] = useState(false);
    const inputRef = useRef(null);
    const [dragActive, setDragActive] = useState(false);

    // --- File Parsing Logic (Updated for Async) ---

    const parseActiveFile = useCallback((file, index) => {
        if (!file) return;
        setFileName(file.name);
        setStatus('Reading file...');

        const isCsv = /\.csv$/i.test(file.name);
        const isExcel = /\.xlsx?$/i.test(file.name);

        if (!isCsv && !isExcel) {
            setStatus('Unsupported file type. Please select .csv or .xlsx');
            return;
        }

        const reader = new FileReader();
        reader.onerror = () => setStatus('Failed to read file.');
        
        reader.onload = async (e) => {
            try {
                let data = [];
                let extractedHeaders = [];

                if (isCsv) {
                    const text = e.target.result;
                    data = parseCSV(text);
                } else {
                    // Async Excel Parse
                    const buffer = e.target.result;
                    data = await parseExcel(buffer);
                }

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

                try {
                    const info = getImportType(extractedHeaders);
                    setFileInfos((prev) => {
                        const copy = [...prev];
                        copy[index] = info;
                        return copy;
                    });
                } catch (e) { /* ignore */ }

            } catch (err) {
                setStatus('Error parsing file.');
                console.error(err);
            }
        };

        if (isCsv) {
            reader.readAsText(file);
        } else {
            reader.readAsArrayBuffer(file);
        }
    }, []);

    const handleAddFiles = useCallback((filesList) => {
        if (!filesList || filesList.length === 0) return;
        const filesArr = Array.from(filesList);
        const validExtRegex = /\.(csv|xlsx?)$/i; 

        // Reset scenario
        if (isImported || importCompleted) {
            const seen = new Set();
            const uniques = [];
            filesArr.forEach((f) => {
                if (f && validExtRegex.test(f.name) && !seen.has(f.name)) {
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

            uniques.forEach((f, i) => {
                detectFileInfo(f, (info) => {
                    setFileInfos(prev => {
                        const copy = [...prev];
                        copy[i] = info;
                        return copy;
                    });
                });
            });

            const firstValidIdx = uniques.findIndex((f) => validExtRegex.test(f.name));
            if (firstValidIdx !== -1) {
                parseActiveFile(uniques[firstValidIdx], firstValidIdx);
            }
            return;
        }

        // Append scenario
        setUploadedFiles((prev) => {
            const copy = [...prev];
            let parsedNewFile = false;

            filesArr.forEach((f) => {
                if (!validExtRegex.test(f.name)) return;

                const existingIndex = copy.findIndex((p) => p.name === f.name);
                
                if (existingIndex !== -1) {
                    copy[existingIndex] = f;
                    setFileInfos(prevInfos => {
                        const infos = [...prevInfos];
                        infos[existingIndex] = null;
                        return infos;
                    });

                    detectFileInfo(f, (info) => {
                        setFileInfos(prev2 => {
                            const copy2 = [...prev2];
                            copy2[existingIndex] = info;
                            return copy2;
                        });
                    });
                    parseActiveFile(f, existingIndex);
                } else {
                    const newIndex = copy.length;
                    copy.push(f);
                    setFileInfos(prevInfos => [...prevInfos, null]);

                    detectFileInfo(f, (info) => {
                        setFileInfos(prev2 => {
                            const copy2 = [...prev2];
                            copy2[newIndex] = info;
                            return copy2;
                        });
                    });

                    if (!parsedNewFile) {
                        parsedNewFile = true;
                        parseActiveFile(f, newIndex);
                    }
                }
            });
            return copy;
        });
    }, [isImported, importCompleted, parseActiveFile]);


    // --- Memoized Data Transformations ---

    const columns = headers;
    const importInfo = useMemo(() => getImportType(columns), [Array.isArray(columns) ? columns.join('|') : String(columns)]);
    const effectiveExcludeFilter = useMemo(() => (excludeFilter && excludeFilter.column) ? excludeFilter : (importInfo && importInfo.excludeFilter), [excludeFilter, importInfo]);
    const effectiveHyperLink = useMemo(() => (hyperLink && hyperLink.column) ? hyperLink : (importInfo && importInfo.hyperLink), [hyperLink, importInfo]);
    const renamed = useMemo(() => applyRenames(parsedData, headers, importInfo && importInfo.rename), [parsedData, headers, importInfo]);
    const enriched = useMemo(() => applyHyperlink(renamed, effectiveHyperLink), [renamed, effectiveHyperLink]);
    const filteredData = useMemo(() => applyExclusion(enriched.data, enriched.headers, effectiveExcludeFilter), [enriched, effectiveExcludeFilter]);
    const matchedWithLink = useMemo(() => {
        const base = Array.isArray(importInfo && importInfo.matched) ? [...importInfo.matched] : [];
        const rename = importInfo && importInfo.rename;
        if (rename) Object.values(rename).forEach(target => { if (target && !base.includes(target)) base.push(target); });
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
        const firstIdx = uploadedFiles.findIndex((f) => f && /\.(csv|xlsx?)$/i.test(f.name));
        if (firstIdx === -1) {
            setStatus('No supported files found.');
            return;
        }

        setProcessingIndex(firstIdx);
        setLastEmittedIndex(-1);
        setIsImported(true);
        setImportCompleted(false);
        setStatus(`Starting import 1 of ${uploadedFiles.length}...`);

        parseActiveFile(uploadedFiles[firstIdx], firstIdx);
    };

    const handleProcessorComplete = useCallback((result) => {
        if (result && result.success) {
            setStatus(`Import ${processingIndex + 1} completed.`);
        } else {
            const msg = result?.error || result?.reason || 'Unknown error';
            setStatus(`Import ${processingIndex + 1} failed: ${msg}`);
        }

        const total = uploadedFiles.length;
        let next = processingIndex + 1;
        while (next < total && !(/\.(csv|xlsx?)$/i.test(uploadedFiles[next]?.name))) {
            next++;
        }

        if (next < total) {
            setProcessingIndex(next);
            setLastEmittedIndex(-1);
            parseActiveFile(uploadedFiles[next], next);
        } else {
            setIsImported(false);
            setProcessingIndex(-1);
            setActiveIndex(-1);
            setImportCompleted(true);
            console.log('ImportManager: all files processed.');
            setStatus('All imports completed.');
        }
    }, [uploadedFiles, processingIndex, parseActiveFile]);

    useEffect(() => {
        if (!isImported || processingIndex === -1) return;
        if (activeIndex !== processingIndex) return;
        if (lastEmittedIndex === processingIndex) return;

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

        const payload = {
            file: uploadedFiles[processingIndex],
            data: filteredData,
            type: importInfo.type || 'csv',
            matched: matchedWithLink,
            headers: enriched.headers || renamed.headers,
        };

        if (typeof onImport === 'function') onImport(payload);
        else console.log('Imported data', payload);

        setLastEmittedIndex(processingIndex);
        setStatus(`Processing import ${processingIndex + 1} of ${uploadedFiles.length}...`);
    }, [isImported, processingIndex, activeIndex, headers, filteredData, importInfo, matchedWithLink, enriched, renamed, uploadedFiles, lastEmittedIndex, onImport]);


    // --- Render ---

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
            <div className="mb-6">
                <div>
                    <h2 className="text-2xl text-slate-800 font-bold tracking-tight">Import Data</h2>
                    <p className="text-slate-400 text-sm mt-1">
                        {hasFiles ? 'Review your files below' : 'Upload CSV or Excel files to begin'}
                    </p>
                </div>
                {status && (
                    <div className="mt-3 animate-in fade-in slide-in-from-top-2 duration-300">
                        <div className="inline-block text-xs font-medium px-3 py-1 bg-slate-100 text-slate-500 rounded-full animate-pulse">
                            {status}
                        </div>
                    </div>
                )}
            </div>

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
                            <div className="absolute -top-4 -left-12 bg-white p-2.5 rounded-xl shadow-lg border border-slate-50 transform -rotate-12 transition-transform duration-500 group-hover:-translate-x-2 group-hover:-rotate-12">
                                <FileText size={24} className="text-emerald-500" />
                                <span className="text-[8px] font-bold text-slate-400 absolute bottom-0.5 right-1">CSV</span>
                            </div>
                            <div className="absolute -top-2 -right-10 bg-white p-2.5 rounded-xl shadow-lg border border-slate-50 transform rotate-12 transition-transform duration-500 group-hover:translate-x-2 group-hover:rotate-6">
                                <Table size={24} className="text-green-600" />
                                <span className="text-[8px] font-bold text-slate-400 absolute bottom-0.5 right-1">XLSX</span>
                            </div>
                        </div>
                        <div className="text-center space-y-2 pointer-events-none">
                            <h3 className="text-lg font-semibold text-slate-700">Drag & drop</h3>
                            <p className="text-sm text-slate-400 font-medium">
                                files here (.csv, .xlsx), or <span className="text-indigo-500 hover:text-indigo-600 underline decoration-indigo-200 underline-offset-2">browse</span>
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
                <input ref={inputRef} type="file" accept=".csv, .xlsx, .xls" multiple className="hidden" onChange={(e) => e.target.files?.length && handleAddFiles(e.target.files)} />
            </div>

            <div className="mt-8 border-t border-slate-100 pt-4">
                <button onClick={() => setShowInfo(!showInfo)} className="flex items-center gap-2 text-sm font-medium text-slate-500 hover:text-indigo-600 transition-colors w-full justify-between group outline-none">
                    <span className="flex items-center gap-2">
                        <Info size={16} className="group-hover:text-indigo-500 transition-colors" />
                        Supported Import Types
                    </span>
                    {showInfo ? <ChevronUp size={16} /> : <ChevronDown size={16} />}
                </button>
                <div className={`grid transition-[grid-template-rows,opacity,margin] duration-300 ease-in-out ${showInfo ? 'grid-rows-[1fr] opacity-100 mt-4' : 'grid-rows-[0fr] opacity-0 mt-0'}`}>
                    <div className="overflow-hidden">
                        <div className="grid gap-3">
                            {IMPORT_DEFINITIONS.map((def) => (
                                <div key={def.id} className="flex items-start gap-3 p-3 bg-slate-50/80 rounded-xl border border-slate-100 hover:border-indigo-100 transition-colors">
                                    {def.icon ? <img src={def.icon} alt="" className="w-8 h-8 object-contain opacity-90 mt-1 rounded-lg" /> : <div className="w-8 h-8 bg-white rounded-lg border border-slate-100 flex items-center justify-center mt-1 shadow-sm"><FileText size={16} className="text-slate-400" /></div>}
                                    <div className="min-w-0 flex-1">
                                        <div className="flex justify-between items-start"><h4 className="text-sm font-semibold text-slate-700 truncate">{def.name}</h4></div>
                                        <div className="flex flex-wrap gap-1.5 mt-1.5">
                                             <span className={`text-[10px] font-bold px-2 py-0.5 rounded-full border ${def.action === 'Update' ? 'bg-amber-50 text-amber-700 border-amber-100' : 'bg-emerald-50 text-emerald-700 border-emerald-100'}`}>{def.action} Mode</span>
                                             <span className="text-[10px] font-medium px-2 py-0.5 rounded-full bg-white border border-slate-200 text-slate-500">{def.type}</span>
                                        </div>
                                        <p className="text-xs text-slate-500 mt-2 leading-relaxed"><span className="font-medium text-slate-600">Required Columns:</span> <span className="font-mono text-indigo-500/90">{def.matchColumns.join(', ')}</span></p>
                                    </div>
                                </div>
                            ))}
                        </div>
                    </div>
                </div>
            </div>

            {parsedData && !isImported && !importCompleted && (
                <div className="mt-6 flex justify-center animate-in fade-in slide-in-from-bottom-4 duration-500 fill-mode-forwards">
                    <button type="button" onClick={handleImport} className="group relative flex items-center justify-center gap-3 bg-slate-900 hover:bg-slate-800 text-white px-8 py-3.5 rounded-full shadow-xl shadow-slate-200 hover:shadow-2xl hover:shadow-slate-300 transform transition-all duration-200 hover:-translate-y-0.5 active:translate-y-0 w-full sm:w-auto sm:min-w-[200px]">
                        <span className="font-semibold tracking-wide text-sm">Import Data</span>
                        {ImportIcon ? <img src={ImportIcon} alt="" className="w-5 h-5 object-contain invert opacity-90 group-hover:opacity-100 transition-opacity" /> : <ArrowRight size={18} className="opacity-70 group-hover:translate-x-1 transition-transform" />}
                    </button>
                </div>
            )}

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