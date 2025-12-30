import React, { useState, useEffect } from 'react';
import { MAPPING_OPERATORS } from '../utils/constants';

export default function CustomParamModal({ isOpen, onClose, customParameters, onSave }) {
    const [showManageModal, setShowManageModal] = useState(false);
    const [editingParam, setEditingParam] = useState(null);
    const [paramName, setParamName] = useState('');
    const [sourceColumn, setSourceColumn] = useState('');
    const [logicType, setLogicType] = useState('');
    const [mappings, setMappings] = useState([]);
    const [customScript, setCustomScript] = useState('');
    const [scriptInputs, setScriptInputs] = useState({});
    const [saveStatus, setSaveStatus] = useState('');

    useEffect(() => {
        if (!isOpen) {
            resetForm();
        }
    }, [isOpen]);

    const resetForm = () => {
        setEditingParam(null);
        setParamName('');
        setSourceColumn('');
        setLogicType('');
        setMappings([]);
        setCustomScript('');
        setScriptInputs({});
        setSaveStatus('');
    };

    const handleEditParam = (param) => {
        setEditingParam(param);
        setParamName(param.name);
        setSourceColumn(param.sourceColumn);
        setLogicType(param.logicType || '');
        setMappings(param.mappings || []);
        setCustomScript(param.script || '');
        setScriptInputs(param.scriptInputs || {});
        setShowManageModal(false);
    };

    const handleAddMapping = () => {
        setMappings([...mappings, { if: '', operator: 'eq', then: '' }]);
    };

    const handleRemoveMapping = (index) => {
        setMappings(mappings.filter((_, i) => i !== index));
    };

    const handleMappingChange = (index, field, value) => {
        const newMappings = [...mappings];
        newMappings[index][field] = value;
        setMappings(newMappings);
    };

    const handleScanScript = () => {
        const regex = /\b(?:let|const|var)\s+([a-zA-Z_$][a-zA-Z0-9_$]*)(?!\s*=)/g;
        const newInputs = new Set();
        let match;

        while ((match = regex.exec(customScript)) !== null) {
            if (match[1] !== 'getWorksheet' && match[1] !== 'sourceColumnValue') {
                newInputs.add(match[1]);
            }
        }

        const updatedInputs = {};
        newInputs.forEach(name => {
            updatedInputs[name] = scriptInputs[name] || '';
        });
        setScriptInputs(updatedInputs);
    };

    const handleScriptInputChange = (varName, value) => {
        setScriptInputs({ ...scriptInputs, [varName]: value });
    };

    const handleFileImport = (e) => {
        const file = e.target.files[0];
        if (file) {
            const reader = new FileReader();
            reader.onload = (event) => {
                setCustomScript(event.target.result);
                handleScanScript();
            };
            reader.readAsText(file);
        }
        e.target.value = '';
    };

    const handleSave = async () => {
        if (!/^[a-zA-Z0-9_]+$/.test(paramName)) {
            setSaveStatus('Parameter Name can only contain letters, numbers, and underscores.');
            return;
        }

        const newParam = {
            name: paramName,
            sourceColumn: sourceColumn,
            logicType: logicType
        };

        if (logicType === 'value-mapping') {
            newParam.mappings = mappings.filter(m => m.if);
        }

        if (logicType === 'custom-script') {
            newParam.script = customScript.trim();
            newParam.scriptInputs = scriptInputs;
        }

        let updatedParams = [...customParameters];

        // Check for duplicate names
        if (editingParam && editingParam.name !== paramName) {
            if (updatedParams.some(p => p.name === paramName)) {
                setSaveStatus('A parameter with this name already exists.');
                return;
            }
            updatedParams = updatedParams.filter(p => p.name !== editingParam.name);
        } else if (!editingParam && updatedParams.some(p => p.name === paramName)) {
            setSaveStatus('A parameter with this name already exists.');
            return;
        }

        const existingIndex = updatedParams.findIndex(p => p.name === paramName);
        if (existingIndex > -1) {
            updatedParams[existingIndex] = newParam;
        } else {
            updatedParams.push(newParam);
        }

        await onSave(updatedParams);
        setSaveStatus('Parameter saved!');
        setTimeout(() => {
            setSaveStatus('');
            setShowManageModal(true);
        }, 1000);
    };

    const handleDeleteParam = async (paramName) => {
        const updatedParams = customParameters.filter(p => p.name !== paramName);
        await onSave(updatedParams);
    };

    if (!isOpen) return null;

    // Manage Parameters View
    if (showManageModal) {
        return (
            <div className="fixed inset-0 bg-gray-600 bg-opacity-75 flex items-center justify-center z-50 p-4">
                <div className="bg-white rounded-lg shadow-xl p-6 w-full max-w-3xl my-8">
                    <h3 className="text-lg font-semibold text-gray-800 mb-4">Manage Custom Parameters</h3>
                    <div className="max-h-96 overflow-y-auto border-t border-b py-2">
                        {customParameters.length === 0 ? (
                            <p className="text-center text-gray-500 text-sm">No custom parameters created yet.</p>
                        ) : (
                            customParameters.map(param => (
                                <div key={param.name} className="flex items-center justify-between p-2 my-1 rounded-md hover:bg-gray-50">
                                    <div>
                                        <span className="text-sm font-medium text-gray-800">{`{${param.name}}`}</span>
                                        <span className="text-xs text-gray-500 ml-2">(from: {param.sourceColumn})</span>
                                    </div>
                                    <div className="flex space-x-2">
                                        <button
                                            onClick={() => handleEditParam(param)}
                                            className="px-2 py-1 bg-blue-100 text-blue-800 text-xs font-semibold rounded-md hover:bg-blue-200"
                                        >
                                            Edit
                                        </button>
                                        <button
                                            onClick={() => handleDeleteParam(param.name)}
                                            className="px-2 py-1 bg-red-100 text-red-800 text-xs font-semibold rounded-md hover:bg-red-200"
                                        >
                                            Delete
                                        </button>
                                    </div>
                                </div>
                            ))
                        )}
                    </div>
                    <div className="flex justify-end mt-4">
                        <button
                            onClick={() => {
                                setShowManageModal(false);
                                onClose();
                            }}
                            className="px-4 py-2 bg-gray-200 text-gray-700 rounded-md hover:bg-gray-300"
                        >
                            Close
                        </button>
                    </div>
                </div>
            </div>
        );
    }

    // Create/Edit Parameter View
    return (
        <div className="fixed inset-0 bg-gray-600 bg-opacity-75 flex items-center justify-center z-50 overflow-y-auto p-4">
            <div className="bg-white rounded-lg shadow-xl p-6 w-full max-w-2xl my-8">
                <div className="flex justify-between items-center mb-4">
                    <h3 className="text-lg font-semibold text-gray-800">
                        {editingParam ? 'Edit Custom Parameter' : 'Create Custom Parameter'}
                    </h3>
                    <button
                        onClick={() => setShowManageModal(true)}
                        className="text-xs text-blue-600 hover:underline"
                    >
                        Manage Parameters
                    </button>
                </div>

                <div className="space-y-4">
                    <div>
                        <label htmlFor="param-name" className="block text-sm font-medium text-gray-700">
                            Parameter Name
                        </label>
                        <input
                            type="text"
                            id="param-name"
                            value={paramName}
                            onChange={(e) => setParamName(e.target.value)}
                            className="mt-1 block w-full px-3 py-2 border border-gray-300 rounded-md shadow-sm"
                            placeholder="e.g., AdvisorEmail"
                        />
                        <p className="mt-1 text-xs text-gray-500">
                            No spaces or special characters. It will be used like {`{YourName}`}.
                        </p>
                    </div>

                    <div className="space-y-4 border-t pt-4">
                        <div>
                            <label htmlFor="param-source-column" className="block text-sm font-medium text-gray-700">
                                Source Column
                            </label>
                            <input
                                type="text"
                                id="param-source-column"
                                value={sourceColumn}
                                onChange={(e) => setSourceColumn(e.target.value)}
                                className="mt-1 block w-full px-3 py-2 border border-gray-300 rounded-md shadow-sm"
                                placeholder="e.g., Status"
                            />
                            <p className="mt-1 text-xs text-gray-500">
                                The column in your sheet that this parameter will read from.
                            </p>
                        </div>

                        <div>
                            <label htmlFor="logic-type-dropdown" className="block text-sm font-medium text-gray-700">
                                Logic Type (Optional)
                            </label>
                            <select
                                id="logic-type-dropdown"
                                value={logicType}
                                onChange={(e) => setLogicType(e.target.value)}
                                className="mt-1 block w-full px-3 py-2 border border-gray-300 bg-white rounded-md shadow-sm focus:outline-none focus:ring-blue-500 focus:border-blue-500"
                            >
                                <option value="">Select a logic type...</option>
                                <option value="value-mapping">Value Mapping</option>
                                <option value="custom-script">Custom Script</option>
                            </select>
                            <p className="mt-1 text-xs text-gray-500">
                                Optionally, transform the value from the source column before inserting it.
                            </p>
                        </div>

                        {/* Value Mapping Logic */}
                        {logicType === 'value-mapping' && (
                            <div>
                                <h4 className="text-sm font-medium text-gray-700 mb-2">Value Mappings</h4>
                                <p className="mt-1 text-xs text-gray-500 mb-2">
                                    If a cell's value matches a condition, the parameter will be replaced by the corresponding 'then' value.
                                    If no conditions match, the original cell value will be used.
                                </p>
                                <div className="space-y-2">
                                    {mappings.map((mapping, index) => (
                                        <div key={index} className="flex items-center gap-2">
                                            <span className="text-sm">If cell</span>
                                            <select
                                                value={mapping.operator}
                                                onChange={(e) => handleMappingChange(index, 'operator', e.target.value)}
                                                className="w-32 px-2 py-1 border border-gray-300 rounded-md text-sm"
                                            >
                                                {MAPPING_OPERATORS.map(op => (
                                                    <option key={op.value} value={op.value}>{op.text}</option>
                                                ))}
                                            </select>
                                            <input
                                                type="text"
                                                value={mapping.if}
                                                onChange={(e) => handleMappingChange(index, 'if', e.target.value)}
                                                className="flex-grow px-2 py-1 border border-gray-300 rounded-md text-sm"
                                                placeholder="Value..."
                                            />
                                            <span className="text-sm">then</span>
                                            <input
                                                type="text"
                                                value={mapping.then}
                                                onChange={(e) => handleMappingChange(index, 'then', e.target.value)}
                                                className="flex-grow px-2 py-1 border border-gray-300 rounded-md text-sm"
                                                placeholder="Result..."
                                            />
                                            <button
                                                onClick={() => handleRemoveMapping(index)}
                                                className="text-red-500 hover:text-red-700 text-xl"
                                            >
                                                ×
                                            </button>
                                        </div>
                                    ))}
                                </div>
                                <button
                                    onClick={handleAddMapping}
                                    className="mt-2 text-xs text-blue-600 hover:underline"
                                >
                                    + Add Mapping
                                </button>
                            </div>
                        )}

                        {/* Custom Script Logic */}
                        {logicType === 'custom-script' && (
                            <div>
                                <h4 className="text-sm font-medium text-gray-700 mb-2">Custom Script</h4>
                                <div className="space-y-3">
                                    <div>
                                        <label className="text-xs text-blue-600 hover:underline cursor-pointer">
                                            + Import Script from File
                                            <input
                                                type="file"
                                                accept=".js,.txt"
                                                onChange={handleFileImport}
                                                className="hidden"
                                            />
                                        </label>
                                    </div>
                                    <textarea
                                        value={customScript}
                                        onChange={(e) => setCustomScript(e.target.value)}
                                        className="w-full h-32 p-2 border rounded-md font-mono text-xs bg-gray-50 focus:bg-white transition-colors"
                                        placeholder="Paste your script here or import from a file..."
                                    />
                                    <div className="flex justify-end">
                                        <button
                                            onClick={handleScanScript}
                                            className="px-3 py-1 bg-gray-200 text-gray-800 text-xs font-semibold rounded-md hover:bg-gray-300"
                                        >
                                            Scan for Inputs
                                        </button>
                                    </div>

                                    {Object.keys(scriptInputs).length > 0 && (
                                        <div className="space-y-2 border-t pt-3 mt-3">
                                            <h5 className="text-xs font-medium text-gray-600">Script Inputs</h5>
                                            <p className="text-xs text-gray-500">
                                                Provide a source column name for each required script input.
                                            </p>
                                            <div className="space-y-2">
                                                {Object.keys(scriptInputs).map(varName => (
                                                    <div key={varName} className="flex items-center gap-2">
                                                        <label className="w-1/3 text-sm font-mono text-gray-600">
                                                            {varName}
                                                        </label>
                                                        <input
                                                            type="text"
                                                            value={scriptInputs[varName]}
                                                            onChange={(e) => handleScriptInputChange(varName, e.target.value)}
                                                            className="flex-grow px-2 py-1 border border-gray-300 rounded-md text-sm"
                                                            placeholder="Source Column Name"
                                                        />
                                                    </div>
                                                ))}
                                            </div>
                                        </div>
                                    )}
                                </div>
                            </div>
                        )}
                    </div>
                </div>

                <p className="text-xs mt-2 h-4 text-center text-red-600">{saveStatus}</p>

                <div className="flex justify-end gap-2 mt-6 border-t pt-4">
                    <button
                        onClick={onClose}
                        className="px-4 py-2 bg-gray-200 text-gray-700 rounded-md hover:bg-gray-300"
                    >
                        Cancel
                    </button>
                    <button
                        onClick={handleSave}
                        className="px-4 py-2 bg-blue-600 text-white rounded-md hover:bg-blue-700"
                    >
                        Save Parameter
                    </button>
                </div>
            </div>
        </div>
    );
}
