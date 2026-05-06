import React, { useState, useEffect } from 'react';
import { CUSTOM_PARAMS_KEY } from '../personalizedEmail/utils/constants';

const conflictKey = (p) => (p?.name || '').trim().toLowerCase();

const sanitizeImportedParam = (p) => {
    if (!p || typeof p !== 'object') return null;
    if (typeof p.name !== 'string' || !p.name.trim()) return null;
    if (!/^[a-zA-Z0-9_]+$/.test(p.name.trim())) return null;
    return {
        name: p.name.trim(),
        sourceColumn: typeof p.sourceColumn === 'string' ? p.sourceColumn : '',
        mappings: Array.isArray(p.mappings)
            ? p.mappings
                  .filter(m => m && typeof m === 'object')
                  .map(m => ({
                      if: typeof m.if === 'string' ? m.if : '',
                      operator: typeof m.operator === 'string' ? m.operator : 'eq',
                      then: typeof m.then === 'string' ? m.then : '',
                  }))
            : [],
    };
};

export default function CustomParamImportExportModal({ isOpen, onClose }) {
    const [activeTab, setActiveTab] = useState('export');
    const [params, setParams] = useState([]);
    const [isLoading, setIsLoading] = useState(false);

    const [selectedExportNames, setSelectedExportNames] = useState(new Set());
    const [exportStatus, setExportStatus] = useState('');

    const [importedParams, setImportedParams] = useState(null);
    const [importedFileName, setImportedFileName] = useState('');
    const [selectedImportIndices, setSelectedImportIndices] = useState(new Set());
    const [overwriteConflicts, setOverwriteConflicts] = useState(new Set());
    const [importError, setImportError] = useState('');
    const [importStatus, setImportStatus] = useState('');

    useEffect(() => {
        if (!isOpen) return;
        setActiveTab('export');
        setSelectedExportNames(new Set());
        setExportStatus('');
        setImportedParams(null);
        setImportedFileName('');
        setSelectedImportIndices(new Set());
        setOverwriteConflicts(new Set());
        setImportError('');
        setImportStatus('');
        loadParams();
        // eslint-disable-next-line react-hooks/exhaustive-deps
    }, [isOpen]);

    const loadParams = async () => {
        setIsLoading(true);
        try {
            const loaded = await Excel.run(async (context) => {
                const settings = context.workbook.settings;
                const ps = settings.getItemOrNullObject(CUSTOM_PARAMS_KEY);
                ps.load('value');
                await context.sync();
                return ps.value ? JSON.parse(ps.value) : [];
            });
            setParams(Array.isArray(loaded) ? loaded : []);
        } catch (err) {
            console.error('Failed to load custom parameters:', err);
            setParams([]);
        } finally {
            setIsLoading(false);
        }
    };

    const persistParams = async (arr) => {
        await Excel.run(async (context) => {
            context.workbook.settings.add(CUSTOM_PARAMS_KEY, JSON.stringify(arr));
            await context.sync();
        });
    };

    // ---------- Export ----------
    const allExportSelected =
        params.length > 0 && selectedExportNames.size === params.length;

    const toggleExportSelection = (name) => {
        setSelectedExportNames(prev => {
            const next = new Set(prev);
            if (next.has(name)) next.delete(name);
            else next.add(name);
            return next;
        });
    };

    const toggleSelectAllExport = () => {
        if (allExportSelected) setSelectedExportNames(new Set());
        else setSelectedExportNames(new Set(params.map(p => p.name)));
    };

    const handleExport = () => {
        const toExport = params.filter(p => selectedExportNames.has(p.name));
        if (toExport.length === 0) {
            setExportStatus('Select at least one parameter to export.');
            return;
        }
        const payload = {
            type: 'student-retention-custom-parameters',
            version: 1,
            exportedAt: new Date().toISOString(),
            parameters: toExport,
        };
        try {
            const blob = new Blob([JSON.stringify(payload, null, 2)], {
                type: 'application/json',
            });
            const url = URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = `custom-parameters-${new Date().toISOString().slice(0, 10)}.json`;
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
            URL.revokeObjectURL(url);
            setExportStatus(`Exported ${toExport.length} parameter${toExport.length === 1 ? '' : 's'}.`);
        } catch (err) {
            console.error('Export failed:', err);
            setExportStatus('Export failed. See console for details.');
        }
    };

    // ---------- Import ----------
    const findConflict = (param) => {
        if (!param) return null;
        const key = conflictKey(param);
        return params.find(existing => conflictKey(existing) === key) || null;
    };

    const handleFileSelect = (e) => {
        const file = e.target.files && e.target.files[0];
        if (!file) return;
        setImportError('');
        setImportStatus('');
        const reader = new FileReader();
        reader.onload = () => {
            try {
                const parsed = JSON.parse(reader.result);
                let candidates;
                if (Array.isArray(parsed)) {
                    candidates = parsed;
                } else if (parsed && Array.isArray(parsed.parameters)) {
                    candidates = parsed.parameters;
                } else {
                    throw new Error('Invalid file: expected an array or an object with a "parameters" array.');
                }
                const valid = candidates
                    .map(p => sanitizeImportedParam(p))
                    .filter(Boolean);
                if (valid.length === 0) {
                    throw new Error('No valid parameters found in file.');
                }
                setImportedParams(valid);
                setImportedFileName(file.name);
                setSelectedImportIndices(new Set(valid.map((_, i) => i)));
                setOverwriteConflicts(new Set());
                setImportStatus(`Loaded ${valid.length} parameter${valid.length === 1 ? '' : 's'} from "${file.name}".`);
            } catch (err) {
                setImportedParams(null);
                setImportedFileName('');
                setSelectedImportIndices(new Set());
                setOverwriteConflicts(new Set());
                setImportError(err.message || 'Failed to parse JSON file.');
            }
        };
        reader.onerror = () => {
            setImportError('Failed to read file.');
            setImportedParams(null);
        };
        reader.readAsText(file);
        e.target.value = '';
    };

    const allImportSelected =
        !!importedParams && selectedImportIndices.size === importedParams.length;

    const toggleImportSelection = (index) => {
        setSelectedImportIndices(prev => {
            const next = new Set(prev);
            if (next.has(index)) next.delete(index);
            else next.add(index);
            return next;
        });
    };

    const toggleSelectAllImport = () => {
        if (!importedParams) return;
        if (allImportSelected) setSelectedImportIndices(new Set());
        else setSelectedImportIndices(new Set(importedParams.map((_, i) => i)));
    };

    const toggleOverwrite = (index) => {
        setOverwriteConflicts(prev => {
            const next = new Set(prev);
            if (next.has(index)) next.delete(index);
            else next.add(index);
            return next;
        });
    };

    const conflictIndices = importedParams
        ? importedParams
              .map((p, i) => (findConflict(p) ? i : -1))
              .filter(i => i !== -1)
        : [];

    const allConflictsOverwritten =
        conflictIndices.length > 0 &&
        conflictIndices.every(i => overwriteConflicts.has(i));

    const toggleOverwriteAllConflicts = () => {
        if (allConflictsOverwritten) {
            setOverwriteConflicts(new Set());
        } else {
            setOverwriteConflicts(new Set(conflictIndices));
        }
    };

    const handleImport = async () => {
        if (!importedParams) return;
        const selectedCount = selectedImportIndices.size;
        if (selectedCount === 0) {
            setImportStatus('Select at least one parameter to import.');
            return;
        }
        setIsLoading(true);
        try {
            const merged = [...params];
            let added = 0;
            let overwritten = 0;
            let skipped = 0;

            importedParams.forEach((param, i) => {
                if (!selectedImportIndices.has(i)) return;
                const conflict = findConflict(param);
                if (conflict) {
                    if (overwriteConflicts.has(i)) {
                        const idx = merged.findIndex(p => conflictKey(p) === conflictKey(conflict));
                        if (idx !== -1) {
                            merged[idx] = {
                                name: param.name,
                                sourceColumn: param.sourceColumn,
                                mappings: param.mappings,
                            };
                            overwritten++;
                        }
                    } else {
                        skipped++;
                    }
                } else {
                    merged.push({
                        name: param.name,
                        sourceColumn: param.sourceColumn,
                        mappings: param.mappings,
                    });
                    added++;
                }
            });

            await persistParams(merged);
            setParams(merged);
            setImportedParams(null);
            setImportedFileName('');
            setSelectedImportIndices(new Set());
            setOverwriteConflicts(new Set());

            const parts = [];
            if (added) parts.push(`${added} added`);
            if (overwritten) parts.push(`${overwritten} overwritten`);
            if (skipped) parts.push(`${skipped} skipped (conflict)`);
            setImportStatus(`Import complete: ${parts.join(', ') || 'no changes'}.`);
        } catch (err) {
            console.error('Failed to import:', err);
            setImportError(err.message || 'Import failed.');
        } finally {
            setIsLoading(false);
        }
    };

    if (!isOpen) return null;

    return (
        <div className="fixed inset-0 bg-gray-600 bg-opacity-75 flex items-center justify-center z-50">
            <div className="bg-white rounded-lg shadow-xl p-6 w-full max-w-2xl">
                <div className="flex justify-between items-start mb-4">
                    <div>
                        <h3 className="text-lg font-semibold text-gray-800">Custom Parameters</h3>
                        <p className="text-xs text-gray-500 mt-0.5">Import or export custom parameters as JSON.</p>
                    </div>
                    <button
                        onClick={onClose}
                        aria-label="Close"
                        className="text-gray-400 hover:text-gray-600 text-xl leading-none"
                    >
                        &times;
                    </button>
                </div>

                <div className="flex border-b border-gray-200 mb-4">
                    <button
                        type="button"
                        className={`px-4 py-2 text-sm font-medium ${
                            activeTab === 'export'
                                ? 'border-b-2 border-blue-600 text-blue-600'
                                : 'text-gray-500 hover:text-gray-700'
                        }`}
                        onClick={() => setActiveTab('export')}
                    >
                        Export
                    </button>
                    <button
                        type="button"
                        className={`px-4 py-2 text-sm font-medium ${
                            activeTab === 'import'
                                ? 'border-b-2 border-blue-600 text-blue-600'
                                : 'text-gray-500 hover:text-gray-700'
                        }`}
                        onClick={() => setActiveTab('import')}
                    >
                        Import
                    </button>
                </div>

                {activeTab === 'export' && (
                    <div>
                        <p className="text-sm text-gray-600 mb-3">
                            Choose which parameters to export to a JSON file.
                        </p>
                        {isLoading && params.length === 0 ? (
                            <p className="text-sm text-gray-500 italic text-center py-6">Loading…</p>
                        ) : params.length === 0 ? (
                            <p className="text-sm text-gray-500 italic text-center py-6">
                                No saved custom parameters to export.
                            </p>
                        ) : (
                            <>
                                <div className="flex items-center justify-between border-b pb-2 mb-2">
                                    <button
                                        type="button"
                                        onClick={toggleSelectAllExport}
                                        className="text-xs font-medium text-blue-600 hover:text-blue-800"
                                    >
                                        {allExportSelected ? 'Deselect All' : 'Select All'}
                                    </button>
                                    <span className="text-xs text-gray-500">
                                        {selectedExportNames.size} of {params.length} selected
                                    </span>
                                </div>
                                <div className="space-y-1 max-h-72 overflow-y-auto">
                                    {params.map(p => (
                                        <label
                                            key={p.name}
                                            className="flex items-center gap-3 p-2 rounded hover:bg-gray-50 cursor-pointer"
                                        >
                                            <input
                                                type="checkbox"
                                                checked={selectedExportNames.has(p.name)}
                                                onChange={() => toggleExportSelection(p.name)}
                                            />
                                            <div className="min-w-0 flex-1">
                                                <div className="text-sm font-medium text-gray-800 truncate">
                                                    {`{${p.name}}`}
                                                </div>
                                                <div className="text-xs text-gray-500 truncate">
                                                    from: {p.sourceColumn || '—'}
                                                    {Array.isArray(p.mappings) && p.mappings.length > 0
                                                        ? ` · ${p.mappings.length} mapping${p.mappings.length === 1 ? '' : 's'}`
                                                        : ''}
                                                </div>
                                            </div>
                                        </label>
                                    ))}
                                </div>
                            </>
                        )}
                        {exportStatus && (
                            <div className="text-sm text-gray-700 mt-3">{exportStatus}</div>
                        )}
                        <div className="flex justify-end gap-2 mt-4">
                            <button
                                onClick={onClose}
                                className="px-4 py-2 bg-gray-200 text-gray-700 rounded-md hover:bg-gray-300"
                            >
                                Close
                            </button>
                            <button
                                onClick={handleExport}
                                disabled={selectedExportNames.size === 0}
                                className={`px-4 py-2 text-white rounded-md ${
                                    selectedExportNames.size === 0
                                        ? 'bg-blue-300 cursor-not-allowed'
                                        : 'bg-blue-600 hover:bg-blue-700'
                                }`}
                            >
                                Export ({selectedExportNames.size})
                            </button>
                        </div>
                    </div>
                )}

                {activeTab === 'import' && (
                    <div>
                        <p className="text-sm text-gray-600 mb-3">
                            Choose a JSON file. Conflicts (same name) are flagged so you can choose to overwrite or skip.
                        </p>
                        <div className="flex items-center gap-3 mb-3">
                            <label className="inline-block px-4 py-2 bg-blue-600 text-white rounded-md hover:bg-blue-700 cursor-pointer text-sm">
                                Choose File
                                <input
                                    type="file"
                                    accept=".json,application/json"
                                    onChange={handleFileSelect}
                                    className="hidden"
                                />
                            </label>
                            {importedFileName && (
                                <span className="text-xs text-gray-500 truncate">{importedFileName}</span>
                            )}
                        </div>
                        {importError && (
                            <div className="text-sm text-red-600 mb-2">{importError}</div>
                        )}
                        {importedParams && importedParams.length > 0 && (
                            <>
                                <div className="flex items-center justify-between border-b pb-2 mb-2">
                                    <div className="flex gap-3">
                                        <button
                                            type="button"
                                            onClick={toggleSelectAllImport}
                                            className="text-xs font-medium text-blue-600 hover:text-blue-800"
                                        >
                                            {allImportSelected ? 'Deselect All' : 'Select All'}
                                        </button>
                                        {conflictIndices.length > 0 && (
                                            <button
                                                type="button"
                                                onClick={toggleOverwriteAllConflicts}
                                                className="text-xs font-medium text-amber-700 hover:text-amber-900"
                                            >
                                                {allConflictsOverwritten
                                                    ? 'Skip All Conflicts'
                                                    : `Overwrite All Conflicts (${conflictIndices.length})`}
                                            </button>
                                        )}
                                    </div>
                                    <span className="text-xs text-gray-500">
                                        {selectedImportIndices.size} of {importedParams.length} selected
                                        {conflictIndices.length > 0 && ` · ${conflictIndices.length} conflict${conflictIndices.length === 1 ? '' : 's'}`}
                                    </span>
                                </div>
                                <div className="space-y-1 max-h-72 overflow-y-auto">
                                    {importedParams.map((p, i) => {
                                        const conflict = findConflict(p);
                                        const isSelected = selectedImportIndices.has(i);
                                        return (
                                            <div
                                                key={i}
                                                className={`flex items-start justify-between gap-3 p-2 rounded ${
                                                    conflict
                                                        ? 'bg-amber-50 border border-amber-200'
                                                        : 'hover:bg-gray-50'
                                                }`}
                                            >
                                                <label className="flex items-start gap-3 min-w-0 flex-1 cursor-pointer">
                                                    <input
                                                        type="checkbox"
                                                        checked={isSelected}
                                                        onChange={() => toggleImportSelection(i)}
                                                        className="mt-1"
                                                    />
                                                    <div className="min-w-0 flex-1">
                                                        <div className="text-sm font-medium text-gray-800 truncate">
                                                            {`{${p.name}}`}
                                                        </div>
                                                        <div className="text-xs text-gray-500 truncate">
                                                            from: {p.sourceColumn || '—'}
                                                            {Array.isArray(p.mappings) && p.mappings.length > 0
                                                                ? ` · ${p.mappings.length} mapping${p.mappings.length === 1 ? '' : 's'}`
                                                                : ''}
                                                        </div>
                                                        {conflict && (
                                                            <div className="text-xs text-amber-700 mt-1">
                                                                Conflict: a parameter with this name already exists.
                                                            </div>
                                                        )}
                                                    </div>
                                                </label>
                                                {conflict && isSelected && (
                                                    <label className="flex items-center gap-1 text-xs text-amber-800 cursor-pointer whitespace-nowrap">
                                                        <input
                                                            type="checkbox"
                                                            checked={overwriteConflicts.has(i)}
                                                            onChange={() => toggleOverwrite(i)}
                                                        />
                                                        Overwrite
                                                    </label>
                                                )}
                                            </div>
                                        );
                                    })}
                                </div>
                            </>
                        )}
                        {importStatus && !importError && (
                            <div className="text-sm text-gray-700 mt-3">{importStatus}</div>
                        )}
                        <div className="flex justify-end gap-2 mt-4">
                            <button
                                onClick={onClose}
                                className="px-4 py-2 bg-gray-200 text-gray-700 rounded-md hover:bg-gray-300"
                            >
                                Close
                            </button>
                            <button
                                onClick={handleImport}
                                disabled={
                                    !importedParams ||
                                    selectedImportIndices.size === 0 ||
                                    isLoading
                                }
                                className={`px-4 py-2 text-white rounded-md ${
                                    !importedParams ||
                                    selectedImportIndices.size === 0 ||
                                    isLoading
                                        ? 'bg-blue-300 cursor-not-allowed'
                                        : 'bg-blue-600 hover:bg-blue-700'
                                }`}
                            >
                                {isLoading
                                    ? 'Importing…'
                                    : `Import (${importedParams ? selectedImportIndices.size : 0})`}
                            </button>
                        </div>
                    </div>
                )}
            </div>
        </div>
    );
}
