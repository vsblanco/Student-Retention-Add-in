import React, { useState, useEffect } from 'react';
import { EMAIL_TEMPLATES_KEY } from '../personalizedEmail/utils/constants';

const conflictKey = (t) =>
    `${(t?.name || '').trim().toLowerCase()}::${(t?.author || '').trim().toLowerCase()}`;

const sanitizeImportedTemplate = (t, index) => {
    if (!t || typeof t !== 'object') return null;
    if (typeof t.name !== 'string' || !t.name.trim()) return null;
    return {
        // Keep the original id if it exists; new ids are minted on insert.
        id: typeof t.id === 'string' ? t.id : `import-tpl-${Date.now()}-${index}`,
        name: t.name.trim(),
        author: typeof t.author === 'string' && t.author.trim() ? t.author.trim() : 'Imported',
        from: typeof t.from === 'string' ? t.from : '',
        subject: typeof t.subject === 'string' ? t.subject : '',
        body: typeof t.body === 'string' ? t.body : '',
        cc: Array.isArray(t.cc) ? t.cc.filter(x => typeof x === 'string') : [],
        createdAt: typeof t.createdAt === 'string' ? t.createdAt : new Date().toISOString(),
    };
};

export default function TemplateImportExportModal({ isOpen, onClose }) {
    const [activeTab, setActiveTab] = useState('export');
    const [templates, setTemplates] = useState([]);
    const [isLoading, setIsLoading] = useState(false);

    const [selectedExportIds, setSelectedExportIds] = useState(new Set());
    const [exportStatus, setExportStatus] = useState('');

    const [importedTemplates, setImportedTemplates] = useState(null);
    const [importedFileName, setImportedFileName] = useState('');
    const [selectedImportIndices, setSelectedImportIndices] = useState(new Set());
    const [overwriteConflicts, setOverwriteConflicts] = useState(new Set());
    const [importError, setImportError] = useState('');
    const [importStatus, setImportStatus] = useState('');

    useEffect(() => {
        if (!isOpen) return;
        setActiveTab('export');
        setSelectedExportIds(new Set());
        setExportStatus('');
        setImportedTemplates(null);
        setImportedFileName('');
        setSelectedImportIndices(new Set());
        setOverwriteConflicts(new Set());
        setImportError('');
        setImportStatus('');
        loadTemplates();
        // eslint-disable-next-line react-hooks/exhaustive-deps
    }, [isOpen]);

    const loadTemplates = async () => {
        setIsLoading(true);
        try {
            const loaded = await Excel.run(async (context) => {
                const settings = context.workbook.settings;
                const ts = settings.getItemOrNullObject(EMAIL_TEMPLATES_KEY);
                ts.load('value');
                await context.sync();
                return ts.value ? JSON.parse(ts.value) : [];
            });
            setTemplates(Array.isArray(loaded) ? loaded : []);
        } catch (err) {
            console.error('Failed to load templates:', err);
            setTemplates([]);
        } finally {
            setIsLoading(false);
        }
    };

    const persistTemplates = async (arr) => {
        await Excel.run(async (context) => {
            context.workbook.settings.add(EMAIL_TEMPLATES_KEY, JSON.stringify(arr));
            await context.sync();
        });
    };

    // ---------- Export ----------
    const allExportSelected =
        templates.length > 0 && selectedExportIds.size === templates.length;

    const toggleExportSelection = (id) => {
        setSelectedExportIds(prev => {
            const next = new Set(prev);
            if (next.has(id)) next.delete(id);
            else next.add(id);
            return next;
        });
    };

    const toggleSelectAllExport = () => {
        if (allExportSelected) setSelectedExportIds(new Set());
        else setSelectedExportIds(new Set(templates.map(t => t.id)));
    };

    const handleExport = () => {
        const toExport = templates.filter(t => selectedExportIds.has(t.id));
        if (toExport.length === 0) {
            setExportStatus('Select at least one template to export.');
            return;
        }
        const payload = {
            type: 'student-retention-email-templates',
            version: 1,
            exportedAt: new Date().toISOString(),
            templates: toExport,
        };
        try {
            const blob = new Blob([JSON.stringify(payload, null, 2)], {
                type: 'application/json',
            });
            const url = URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = `email-templates-${new Date().toISOString().slice(0, 10)}.json`;
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
            URL.revokeObjectURL(url);
            setExportStatus(`Exported ${toExport.length} template${toExport.length === 1 ? '' : 's'}.`);
        } catch (err) {
            console.error('Export failed:', err);
            setExportStatus('Export failed. See console for details.');
        }
    };

    // ---------- Import ----------
    const findConflict = (template) => {
        if (!template) return null;
        const key = conflictKey(template);
        return templates.find(existing => conflictKey(existing) === key) || null;
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
                } else if (parsed && Array.isArray(parsed.templates)) {
                    candidates = parsed.templates;
                } else {
                    throw new Error('Invalid file: expected an array or an object with a "templates" array.');
                }
                const valid = candidates
                    .map((t, i) => sanitizeImportedTemplate(t, i))
                    .filter(Boolean);
                if (valid.length === 0) {
                    throw new Error('No valid templates found in file.');
                }
                setImportedTemplates(valid);
                setImportedFileName(file.name);
                setSelectedImportIndices(new Set(valid.map((_, i) => i)));
                setOverwriteConflicts(new Set());
                setImportStatus(`Loaded ${valid.length} template${valid.length === 1 ? '' : 's'} from "${file.name}".`);
            } catch (err) {
                setImportedTemplates(null);
                setImportedFileName('');
                setSelectedImportIndices(new Set());
                setOverwriteConflicts(new Set());
                setImportError(err.message || 'Failed to parse JSON file.');
            }
        };
        reader.onerror = () => {
            setImportError('Failed to read file.');
            setImportedTemplates(null);
        };
        reader.readAsText(file);
        e.target.value = '';
    };

    const allImportSelected =
        !!importedTemplates && selectedImportIndices.size === importedTemplates.length;

    const toggleImportSelection = (index) => {
        setSelectedImportIndices(prev => {
            const next = new Set(prev);
            if (next.has(index)) next.delete(index);
            else next.add(index);
            return next;
        });
    };

    const toggleSelectAllImport = () => {
        if (!importedTemplates) return;
        if (allImportSelected) setSelectedImportIndices(new Set());
        else setSelectedImportIndices(new Set(importedTemplates.map((_, i) => i)));
    };

    const toggleOverwrite = (index) => {
        setOverwriteConflicts(prev => {
            const next = new Set(prev);
            if (next.has(index)) next.delete(index);
            else next.add(index);
            return next;
        });
    };

    const conflictIndices = importedTemplates
        ? importedTemplates
              .map((t, i) => (findConflict(t) ? i : -1))
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
        if (!importedTemplates) return;
        const selectedCount = selectedImportIndices.size;
        if (selectedCount === 0) {
            setImportStatus('Select at least one template to import.');
            return;
        }
        setIsLoading(true);
        try {
            const merged = [...templates];
            let added = 0;
            let overwritten = 0;
            let skipped = 0;

            importedTemplates.forEach((tpl, i) => {
                if (!selectedImportIndices.has(i)) return;
                const conflict = findConflict(tpl);
                if (conflict) {
                    if (overwriteConflicts.has(i)) {
                        const idx = merged.findIndex(t => t.id === conflict.id);
                        if (idx !== -1) {
                            merged[idx] = {
                                ...conflict,
                                name: tpl.name,
                                author: tpl.author,
                                from: tpl.from,
                                subject: tpl.subject,
                                body: tpl.body,
                                cc: tpl.cc,
                            };
                            overwritten++;
                        }
                    } else {
                        skipped++;
                    }
                } else {
                    merged.push({
                        ...tpl,
                        id: `tpl-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`,
                        createdAt: new Date().toISOString(),
                    });
                    added++;
                }
            });

            await persistTemplates(merged);
            setTemplates(merged);
            setImportedTemplates(null);
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
                        <h3 className="text-lg font-semibold text-gray-800">Email Templates</h3>
                        <p className="text-xs text-gray-500 mt-0.5">Import or export templates as JSON.</p>
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
                            Choose which templates to export to a JSON file.
                        </p>
                        {isLoading && templates.length === 0 ? (
                            <p className="text-sm text-gray-500 italic text-center py-6">Loading…</p>
                        ) : templates.length === 0 ? (
                            <p className="text-sm text-gray-500 italic text-center py-6">
                                No saved templates to export.
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
                                        {selectedExportIds.size} of {templates.length} selected
                                    </span>
                                </div>
                                <div className="space-y-1 max-h-72 overflow-y-auto">
                                    {templates.map(t => (
                                        <label
                                            key={t.id}
                                            className="flex items-center gap-3 p-2 rounded hover:bg-gray-50 cursor-pointer"
                                        >
                                            <input
                                                type="checkbox"
                                                checked={selectedExportIds.has(t.id)}
                                                onChange={() => toggleExportSelection(t.id)}
                                            />
                                            <div className="min-w-0 flex-1">
                                                <div className="text-sm font-medium text-gray-800 truncate">
                                                    {t.name}
                                                </div>
                                                <div className="text-xs text-gray-500 truncate">
                                                    by {t.author || 'Unknown'}
                                                    {t.subject ? ` · ${t.subject}` : ''}
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
                                disabled={selectedExportIds.size === 0}
                                className={`px-4 py-2 text-white rounded-md ${
                                    selectedExportIds.size === 0
                                        ? 'bg-blue-300 cursor-not-allowed'
                                        : 'bg-blue-600 hover:bg-blue-700'
                                }`}
                            >
                                Export ({selectedExportIds.size})
                            </button>
                        </div>
                    </div>
                )}

                {activeTab === 'import' && (
                    <div>
                        <p className="text-sm text-gray-600 mb-3">
                            Choose a JSON file. Conflicts (same name &amp; author) are flagged so you can choose to overwrite or skip.
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
                        {importedTemplates && importedTemplates.length > 0 && (
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
                                        {selectedImportIndices.size} of {importedTemplates.length} selected
                                        {conflictIndices.length > 0 && ` · ${conflictIndices.length} conflict${conflictIndices.length === 1 ? '' : 's'}`}
                                    </span>
                                </div>
                                <div className="space-y-1 max-h-72 overflow-y-auto">
                                    {importedTemplates.map((t, i) => {
                                        const conflict = findConflict(t);
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
                                                            {t.name}
                                                        </div>
                                                        <div className="text-xs text-gray-500 truncate">
                                                            by {t.author || 'Unknown'}
                                                            {t.subject ? ` · ${t.subject}` : ''}
                                                        </div>
                                                        {conflict && (
                                                            <div className="text-xs text-amber-700 mt-1">
                                                                Conflict: a template with this name &amp; author already exists.
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
                                    !importedTemplates ||
                                    selectedImportIndices.size === 0 ||
                                    isLoading
                                }
                                className={`px-4 py-2 text-white rounded-md ${
                                    !importedTemplates ||
                                    selectedImportIndices.size === 0 ||
                                    isLoading
                                        ? 'bg-blue-300 cursor-not-allowed'
                                        : 'bg-blue-600 hover:bg-blue-700'
                                }`}
                            >
                                {isLoading
                                    ? 'Importing…'
                                    : `Import (${importedTemplates ? selectedImportIndices.size : 0})`}
                            </button>
                        </div>
                    </div>
                )}
            </div>
        </div>
    );
}
