import React, { useState, useEffect, useRef, useCallback } from 'react';
import ReactQuill from 'react-quill-new';
import 'react-quill-new/dist/quill.snow.css';
import PillInput from './components/PillInput';
import { EMAIL_TEMPLATES_KEY, CUSTOM_PARAMS_KEY, standardParameters, QUILL_EDITOR_CONFIG, PARAMETER_BUTTON_STYLES, COLUMN_MAPPINGS } from './utils/constants';
import { findColumnIndex, getTodaysLdaSheetName, getNameParts, isValidEmail, isValidHttpUrl, evaluateMapping, renderTemplate, renderCCTemplate } from './utils/helpers';
import { generatePdfReceipt } from './utils/receiptGenerator';
import ExampleModal from './modals/ExampleModal';
import TemplatesModal from './modals/TemplatesModal';
import CustomParamModal from './modals/CustomParamModal';
import RecipientModal from './modals/RecipientModal';
import ConfirmSendModal from './modals/ConfirmSendModal';
import SuccessModal from './modals/SuccessModal';

export default function PersonalizedEmail({ onReady }) {
    // Connection state
    const [powerAutomateConnection, setPowerAutomateConnection] = useState(null);
    const [isConnected, setIsConnected] = useState(false);
    const [setupUrl, setSetupUrl] = useState('');
    const [setupStatus, setSetupStatus] = useState('');

    // Email composer state
    const [fromPills, setFromPills] = useState([]);
    const [ccPills, setCcPills] = useState([]);
    const [subject, setSubject] = useState('');
    const [body, setBody] = useState('');
    const [status, setStatus] = useState('');

    // Student data state
    const [studentDataCache, setStudentDataCache] = useState([]);
    const [customParameters, setCustomParameters] = useState([]);
    const [recipientSelection, setRecipientSelection] = useState({
        type: 'lda',
        customSheetName: '',
        excludeDNC: true,
        excludeFillColor: true,
        hasBeenSet: false
    });
    const [recipientDataCache, setRecipientDataCache] = useState(new Map());
    const [worksheetDataCache, setWorksheetDataCache] = useState({});

    // UI state
    const [lastFocusedInput, setLastFocusedInput] = useState(null);
    const [showMoreParams, setShowMoreParams] = useState(false);
    const quillRef = useRef(null);

    // Modal states
    const [showExampleModal, setShowExampleModal] = useState(false);
    const [showTemplatesModal, setShowTemplatesModal] = useState(false);
    const [showCustomParamModal, setShowCustomParamModal] = useState(false);
    const [showRecipientModal, setShowRecipientModal] = useState(false);
    const [showConfirmModal, setShowConfirmModal] = useState(false);
    const [showSuccessModal, setShowSuccessModal] = useState(false);
    const [lastSentPayload, setLastSentPayload] = useState([]);

    // Check for existing connection on mount
    useEffect(() => {
        checkConnection();
        loadCustomParameters();
        // Call onReady to signal that component is loaded
        if (onReady) onReady();
    }, [onReady]);

    // Pre-cache recipient data after connection is established
    useEffect(() => {
        if (isConnected) {
            preCacheRecipientData();
        }
    }, [isConnected]);

    const checkConnection = async () => {
        await Excel.run(async (context) => {
            const settings = context.workbook.settings;
            const connectionsSetting = settings.getItemOrNullObject("connections");
            connectionsSetting.load("value");
            await context.sync();

            const connections = connectionsSetting.value ? JSON.parse(connectionsSetting.value) : [];
            const connection = connections.find(c => c.type === 'power-automate' && c.name === 'Send Personalized Email');

            if (connection) {
                setPowerAutomateConnection(connection);
                setIsConnected(true);
            } else {
                setIsConnected(false);
            }
        });
    };

    const createConnection = async () => {
        if (!isValidHttpUrl(setupUrl)) {
            setSetupStatus("Please enter a valid HTTP URL.");
            return;
        }

        setSetupStatus("Creating connection...");

        await Excel.run(async (context) => {
            const settings = context.workbook.settings;
            const connectionsSetting = settings.getItemOrNullObject("connections");
            connectionsSetting.load("value");
            await context.sync();

            let connections = connectionsSetting.value ? JSON.parse(connectionsSetting.value) : [];
            const newConnection = {
                id: 'pa-' + Math.random().toString(36).substr(2, 9),
                name: 'Send Personalized Email',
                type: 'power-automate',
                url: setupUrl,
                actions: [],
                history: []
            };
            connections.push(newConnection);
            settings.add("connections", JSON.stringify(connections));
            await context.sync();

            setSetupStatus("Connection created successfully!");
            setTimeout(() => {
                checkConnection();
            }, 1500);
        });
    };

    const loadCustomParameters = async () => {
        const params = await getCustomParameters();
        setCustomParameters(params);
    };

    const getCustomParameters = async () => {
        return Excel.run(async (context) => {
            const settings = context.workbook.settings;
            const paramsSetting = settings.getItemOrNullObject(CUSTOM_PARAMS_KEY);
            paramsSetting.load("value");
            await context.sync();
            return paramsSetting.value ? JSON.parse(paramsSetting.value) : [];
        });
    };

    const saveCustomParameters = async (params) => {
        await Excel.run(async (context) => {
            context.workbook.settings.add(CUSTOM_PARAMS_KEY, JSON.stringify(params));
            await context.sync();
        });
        setCustomParameters(params);
    };

    const getWorksheetData = async (sheetNameToFetch) => {
        if (worksheetDataCache[sheetNameToFetch]) return worksheetDataCache[sheetNameToFetch];

        try {
            await Excel.run(async (context) => {
                const sheet = context.workbook.worksheets.getItem(sheetNameToFetch);
                const range = sheet.getUsedRange();
                range.load("values");
                await context.sync();

                const values = range.values;
                const data = (values.length > 1)
                    ? { headers: values[0].map(h => String(h ?? '').toLowerCase()), values: values.slice(1) }
                    : { headers: [], values: [] };

                setWorksheetDataCache(prev => ({ ...prev, [sheetNameToFetch]: data }));
            });
            return worksheetDataCache[sheetNameToFetch];
        } catch (error) {
            if (error.code !== 'ItemNotFound') {
                console.error(`Error loading '${sheetNameToFetch}' sheet:`, error);
            }
            setWorksheetDataCache(prev => ({ ...prev, [sheetNameToFetch]: null }));
            return null;
        }
    };

    const getStudentDataCore = async (selection) => {
        const { type, customSheetName, excludeDNC, excludeFillColor } = selection;
        let sheetName;

        if (type === 'custom') {
            sheetName = customSheetName.trim();
            if (!sheetName) {
                const err = new Error('Custom sheet name is required.');
                err.userFacing = true;
                throw err;
            }
        } else {
            sheetName = type === 'lda' ? getTodaysLdaSheetName() : 'Master List';
        }

        const includedStudents = [];
        const excludedStudents = [];
        setWorksheetDataCache({});

        try {
            await Excel.run(async (context) => {
                const dncStudentIdentifiers = new Set();

                if (excludeDNC) {
                    try {
                        const historySheet = context.workbook.worksheets.getItem("Student History");
                        const historyRange = historySheet.getUsedRange();
                        historyRange.load("values");
                        await context.sync();

                        const historyValues = historyRange.values;
                        if (historyValues.length > 1) {
                            const historyHeaders = historyValues[0].map(h => String(h ?? '').toLowerCase());
                            const identifierIndex = findColumnIndex(historyHeaders, COLUMN_MAPPINGS.StudentIdentifier);
                            const tagsIndex = findColumnIndex(historyHeaders, COLUMN_MAPPINGS.Tags);

                            if (identifierIndex !== -1 && tagsIndex !== -1) {
                                for (let i = 1; i < historyValues.length; i++) {
                                    const row = historyValues[i];
                                    const tagsString = String(row[tagsIndex] || '').toUpperCase();
                                    const individualTags = tagsString.split(',').map(t => t.trim());
                                    const hasExcludableDnc = individualTags.some(tag =>
                                        tag.includes('DNC') && !['DNC - PHONE', 'DNC - OTHER PHONE'].includes(tag)
                                    );

                                    if (hasExcludableDnc) {
                                        const studentIdentifier = row[identifierIndex];
                                        if (studentIdentifier) dncStudentIdentifiers.add(String(studentIdentifier));
                                    }
                                }
                            }
                        }
                    } catch (error) {
                        console.error("Could not process 'Student History' sheet for DNC exclusion.", error);
                    }
                }

                const sheet = context.workbook.worksheets.getItem(sheetName);
                const usedRange = sheet.getUsedRange();
                const cellProperties = usedRange.getCellProperties({ format: { fill: { color: true } } });
                usedRange.load("values");
                await context.sync();

                const values = usedRange.values;
                const formats = cellProperties.value;
                const headers = values[0].map(h => String(h ?? '').toLowerCase());

                const colIndices = {};
                for (const key in COLUMN_MAPPINGS) {
                    colIndices[key] = findColumnIndex(headers, COLUMN_MAPPINGS[key]);
                }

                const customParamIndices = {};
                customParameters.forEach(param => {
                    const headerIndex = headers.indexOf(param.sourceColumn.toLowerCase());
                    if (headerIndex !== -1) customParamIndices[param.name] = headerIndex;
                });

                for (let i = 1; i < values.length; i++) {
                    const row = values[i];
                    if (!row) continue;

                    const studentIdentifier = row[colIndices.StudentIdentifier];
                    const studentNameForRow = row[colIndices.StudentName] || `ID: ${studentIdentifier || 'Unknown'}`;
                    const studentEmail = row[colIndices.StudentEmail] ?? '';

                    if (!isValidEmail(studentEmail)) {
                        excludedStudents.push({ name: studentNameForRow, reason: 'Invalid Email' });
                        continue;
                    }

                    if (excludeDNC && colIndices.StudentIdentifier !== -1) {
                        if (studentIdentifier && dncStudentIdentifiers.has(String(studentIdentifier))) {
                            excludedStudents.push({ name: studentNameForRow, reason: 'DNC Tag' });
                            continue;
                        }
                    }

                    if (excludeFillColor && colIndices.Outreach !== -1) {
                        const cellFormat = formats[i]?.[colIndices.Outreach];
                        const cellColor = cellFormat?.format.fill.color;
                        if (cellColor && cellColor !== '#FFFFFF' && cellColor !== '#000000') {
                            excludedStudents.push({ name: studentNameForRow, reason: 'Fill Color' });
                            continue;
                        }
                    }

                    const studentName = row[colIndices.StudentName] ?? '';
                    const nameParts = getNameParts(studentName);
                    const student = {
                        StudentName: studentName,
                        FirstName: nameParts.first,
                        LastName: nameParts.last,
                        StudentEmail: studentEmail,
                        PersonalEmail: row[colIndices.PersonalEmail] ?? '',
                        Grade: row[colIndices.Grade] ?? '',
                        DaysOut: row[colIndices.DaysOut] ?? '',
                        Assigned: row[colIndices.Assigned] ?? ''
                    };

                    for (const param of customParameters) {
                        let value = '';
                        if (param.logicType === 'custom-script' && param.script) {
                            try {
                                const argNames = ['getWorksheet', 'sourceColumnValue'];
                                const argValues = [getWorksheetData, ''];
                                let userScript = param.script;
                                const mainSourceColIndex = headers.indexOf(param.sourceColumn.toLowerCase());
                                if (mainSourceColIndex !== -1) argValues[1] = row[mainSourceColIndex];

                                if (param.scriptInputs) {
                                    for (const varName in param.scriptInputs) {
                                        const sourceColName = param.scriptInputs[varName];
                                        const sourceColIndex = headers.indexOf(sourceColName.toLowerCase());
                                        argNames.push(varName);
                                        argValues.push((sourceColIndex !== -1) ? row[sourceColIndex] : undefined);
                                        userScript = userScript.replace(new RegExp(`\\blet\\s+${varName}\\s*;`, 'g'), '');
                                    }
                                }

                                const isAsync = /\bawait\b/.test(userScript);
                                const hasReturn = /\breturn\b/.test(userScript);
                                let finalScriptBody = isAsync
                                    ? (hasReturn ? userScript : `return (async () => { ${userScript} })();`)
                                    : (hasReturn ? userScript : `return (() => { "use strict"; ${userScript} })();`);
                                const executor = new Function(...argNames, `return (async () => { "use strict"; ${finalScriptBody} })();`);
                                value = await executor(...argValues);
                            } catch (e) {
                                console.error(`Error executing script for parameter "${param.name}":`, e);
                                value = `[SCRIPT ERROR]`;
                            }
                        } else {
                            const colIndex = customParamIndices[param.name];
                            if (colIndex !== undefined) {
                                const cellValue = row[colIndex] ?? '';
                                let mappingFound = false;
                                if (param.mappings && cellValue !== '') {
                                    for (const mapping of param.mappings) {
                                        if (evaluateMapping(cellValue, mapping)) {
                                            value = mapping.then;
                                            mappingFound = true;
                                            break;
                                        }
                                    }
                                }
                                if (!mappingFound) value = cellValue;
                            }
                        }
                        student[param.name] = value;
                    }
                    includedStudents.push(student);
                }
            });
            return { included: includedStudents, excluded: excludedStudents };
        } catch (error) {
            if (error.code === 'ItemNotFound') {
                error.userFacingMessage = `Error: Sheet "${sheetName}" not found.`;
            }
            throw error;
        }
    };

    const getStudentDataWithUI = async () => {
        setStatus('Fetching students...');
        try {
            const result = await getStudentDataCore(recipientSelection);
            setStudentDataCache(result.included);
            setStatus(`Found ${result.included.length} students.`);
            setTimeout(() => setStatus(''), 3000);
            return result.included;
        } catch (error) {
            const message = error.userFacingMessage || (error.userFacing ? error.message : 'An error occurred while fetching data.');
            setStatus(message);
            throw error;
        }
    };

    const preCacheRecipientData = async () => {
        try {
            const ldaSelection = { type: 'lda', customSheetName: '', excludeDNC: true, excludeFillColor: true };
            const ldaResult = await getStudentDataCore(ldaSelection);
            setRecipientDataCache(prev => new Map(prev).set('lda', ldaResult));

            const masterSelection = { type: 'master', customSheetName: '', excludeDNC: true, excludeFillColor: true };
            const masterResult = await getStudentDataCore(masterSelection);
            setRecipientDataCache(prev => new Map(prev).set('master', masterResult));
        } catch (error) {
            console.warn("Pre-caching failed. This may happen if sheets are not yet created.", error);
        }
    };

    const handleRecipientUpdate = (newSelection, count) => {
        setRecipientSelection({ ...newSelection, hasBeenSet: true });
    };

    const insertParameter = (param) => {
        if (lastFocusedInput === 'quill' && quillRef.current) {
            const editor = quillRef.current.getEditor();
            const range = editor.getSelection(true);
            editor.insertText(range.index, param, 'user');
        } else if (lastFocusedInput === 'from') {
            setFromPills([param]);
        } else if (lastFocusedInput === 'cc') {
            setCcPills(prev => [...prev, param]);
        } else if (lastFocusedInput === 'subject') {
            const input = document.getElementById('email-subject');
            if (input) {
                const start = input.selectionStart || 0;
                const end = input.selectionEnd || 0;
                const text = subject;
                const newValue = text.substring(0, start) + param + text.substring(end);
                setSubject(newValue);
                setTimeout(() => {
                    input.focus();
                    input.selectionStart = input.selectionEnd = start + param.length;
                }, 0);
            }
        } else if (quillRef.current) {
            const editor = quillRef.current.getEditor();
            editor.focus();
            editor.insertText(editor.getLength(), param, 'user');
        }
    };

    const generatePayload = () => {
        const fromTemplate = fromPills[0] || '';
        const bodyHtml = body;

        return studentDataCache.map(student => ({
            from: renderTemplate(fromTemplate, student),
            to: student.StudentEmail || '',
            cc: renderCCTemplate(ccPills, student),
            subject: renderTemplate(subject, student),
            body: renderTemplate(bodyHtml, student)
        })).filter(email => email.to && email.from);
    };

    const executeSend = async () => {
        setShowConfirmModal(false);
        setStatus(`Sending ${studentDataCache.length} emails...`);

        const payload = generatePayload();
        setLastSentPayload(payload);

        if (payload.length === 0) {
            setStatus('No students with valid "To" and "From" email addresses found.');
            return;
        }

        try {
            const response = await fetch(powerAutomateConnection.url, {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify(payload)
            });
            if (!response.ok) throw new Error(`HTTP error! status: ${response.status}`);
            setStatus(`Successfully sent ${payload.length} emails!`);
            setShowSuccessModal(true);
        } catch (error) {
            setStatus(`Failed to send emails: ${error.message}`);
            console.error("Error sending emails:", error);
        }
    };

    const isFormValid = () => {
        const from = fromPills[0] || '';
        const isFromValid = from && from.trim() !== '';
        const isSubjectValid = subject && subject.trim() !== '';
        const isBodyValid = body && body.trim() !== '';
        const areRecipientsValid = recipientSelection.hasBeenSet && studentDataCache.length > 0;
        return isFromValid && isSubjectValid && isBodyValid && areRecipientsValid;
    };

    const getValidationMessage = () => {
        const missing = [];
        if (!fromPills[0] || !fromPills[0].trim()) missing.push('From address');
        if (!recipientSelection.hasBeenSet || studentDataCache.length === 0) missing.push('Recipients');
        if (!subject || !subject.trim()) missing.push('Subject');
        if (!body || !body.trim()) missing.push('Body');
        return missing.length > 0 ? `Required: ${missing.join(', ')}.` : '';
    };

    const renderParameterButton = (param) => {
        const isCustom = typeof param === 'object';
        const paramName = isCustom ? param.name : param;

        let buttonClass = PARAMETER_BUTTON_STYLES.standard;
        if (isCustom) {
            const hasMappings = param.mappings && param.mappings.length > 0;
            const hasNested = hasMappings && param.mappings.some(m => /\{(\w+)\}/.test(m.then));

            if (hasNested) buttonClass = PARAMETER_BUTTON_STYLES.nested;
            else if (hasMappings) buttonClass = PARAMETER_BUTTON_STYLES.mapped;
            else if (param.logicType === 'custom-script') buttonClass = PARAMETER_BUTTON_STYLES.script;
            else buttonClass = PARAMETER_BUTTON_STYLES.custom;
        }

        return (
            <button
                key={paramName}
                className={buttonClass}
                onClick={() => insertParameter(`{${paramName}}`)}
            >
                {`{${paramName}}`}
            </button>
        );
    };

    // Setup Wizard View
    if (!isConnected) {
        return (
            <div className="max-w-md mx-auto p-4 bg-gray-50">
                <div className="text-center">
                    <img
                        src="https://vsblanco.github.io/Student-Retention-Add-in/assets/power-automate-icon.png"
                        alt="Power Automate"
                        className="mx-auto h-16 w-16 mb-4"
                    />
                    <h1 className="text-xl font-bold text-gray-800">Setup Required</h1>
                    <p className="text-sm text-gray-600 mt-2">
                        To send personalized emails, you need to connect this add-in to a Power Automate flow.
                    </p>
                </div>

                <div className="mt-6">
                    <label htmlFor="power-automate-url" className="block text-sm font-medium text-gray-700">
                        Power Automate HTTP URL
                    </label>
                    <input
                        type="text"
                        id="power-automate-url"
                        value={setupUrl}
                        onChange={(e) => setSetupUrl(e.target.value)}
                        className="mt-1 block w-full px-3 py-2 border border-gray-300 rounded-md shadow-sm placeholder-gray-400 focus:outline-none focus:ring-blue-500 focus:border-blue-500"
                        placeholder="Paste your flow's URL here"
                    />
                    <p className="mt-2 text-xs text-gray-500">
                        This URL is generated by the "When a HTTP request is received" trigger in your Power Automate flow.
                    </p>
                </div>

                <div className="mt-6">
                    <button
                        onClick={createConnection}
                        className="w-full bg-blue-600 text-white font-bold py-2 px-4 rounded-lg hover:bg-blue-700 focus:outline-none focus:ring-2 focus:ring-blue-500 focus:ring-opacity-50 transition-colors duration-200"
                    >
                        Create Connection
                    </button>
                </div>
                <p className="text-xs mt-2 h-4 text-center text-gray-600">{setupStatus}</p>
            </div>
        );
    }

    // Email Composer View
    return (
        <div className="max-w-md mx-auto p-4 bg-gray-50">
            <div className="flex justify-between items-center mb-4">
                <h1 className="text-xl font-bold text-gray-800">Personalized Email</h1>
                <button
                    onClick={() => setShowTemplatesModal(true)}
                    className="px-3 py-1 bg-gray-200 text-gray-800 text-sm font-semibold rounded-md hover:bg-gray-300"
                >
                    Templates
                </button>
            </div>

            <div className="space-y-4">
                {/* From Field */}
                <div>
                    <label className="block text-sm font-medium text-gray-700">From</label>
                    <PillInput
                        pills={fromPills}
                        onPillsChange={setFromPills}
                        placeholder="sender@example.com"
                        singleValue={true}
                        onFocus={() => setLastFocusedInput('from')}
                    />
                </div>

                {/* Recipient Selection */}
                <div>
                    <label className="block text-sm font-medium text-gray-700">Recipient</label>
                    <button
                        onClick={() => setShowRecipientModal(true)}
                        className={`mt-1 w-full text-left px-3 py-2 border border-gray-300 rounded-md shadow-sm focus:outline-none focus:ring-2 focus:ring-blue-500 ${
                            recipientSelection.hasBeenSet && studentDataCache.length > 0
                                ? 'bg-green-100 text-green-800 font-semibold'
                                : 'bg-white'
                        }`}
                    >
                        {recipientSelection.hasBeenSet && studentDataCache.length > 0
                            ? `${studentDataCache.length} Student${studentDataCache.length !== 1 ? 's' : ''} Selected`
                            : 'Select Students'}
                    </button>
                </div>

                {/* CC Field */}
                <div>
                    <label className="block text-sm font-medium text-gray-700">CC</label>
                    <PillInput
                        pills={ccPills}
                        onPillsChange={setCcPills}
                        placeholder="Add an additional email"
                        onFocus={() => setLastFocusedInput('cc')}
                    />
                </div>

                {/* Subject */}
                <div>
                    <label htmlFor="email-subject" className="block text-sm font-medium text-gray-700">
                        Subject
                    </label>
                    <input
                        type="text"
                        id="email-subject"
                        value={subject}
                        onChange={(e) => setSubject(e.target.value)}
                        onFocus={() => setLastFocusedInput('subject')}
                        className="mt-1 block w-full px-3 py-2 border border-gray-300 rounded-md shadow-sm"
                        placeholder="Email Subject"
                    />
                </div>

                {/* Body (Quill Editor) */}
                <div>
                    <label className="block text-sm font-medium text-gray-700">Body</label>
                    <ReactQuill
                        ref={quillRef}
                        theme="snow"
                        value={body}
                        onChange={setBody}
                        onFocus={() => setLastFocusedInput('quill')}
                        modules={QUILL_EDITOR_CONFIG.modules}
                        className="mt-1 bg-white"
                        style={{ height: '192px', marginBottom: '50px' }}
                    />
                </div>

                {/* Parameters */}
                <div>
                    <label className="block text-sm font-medium text-gray-700">Insert Parameter</label>
                    <div className="mt-2 flex flex-wrap gap-2">
                        {standardParameters.map(param => renderParameterButton(param))}
                    </div>

                    {customParameters.length > 0 && (
                        <div className="mt-3">
                            <label className="block text-xs font-medium text-gray-600 mb-2">Custom Parameters</label>
                            <div className="flex flex-wrap gap-2">
                                {customParameters.slice(0, 5).map(param => renderParameterButton(param))}
                            </div>
                            {customParameters.length > 5 && (
                                <>
                                    {showMoreParams && (
                                        <div className="flex flex-wrap gap-2 mt-2">
                                            {customParameters.slice(5).map(param => renderParameterButton(param))}
                                        </div>
                                    )}
                                    <button
                                        onClick={() => setShowMoreParams(!showMoreParams)}
                                        className="mt-2 text-xs text-blue-600 hover:underline"
                                    >
                                        {showMoreParams ? 'Show Less' : `Show ${customParameters.length - 5} More...`}
                                    </button>
                                </>
                            )}
                        </div>
                    )}

                    <button
                        onClick={() => setShowCustomParamModal(true)}
                        className="mt-2 text-xs text-blue-600 hover:underline"
                    >
                        + Create Custom Parameter
                    </button>
                </div>
            </div>

            {/* Action Buttons */}
            <div className="mt-6 flex space-x-2">
                <button
                    onClick={() => setShowExampleModal(true)}
                    className="w-1/2 bg-gray-200 text-gray-800 font-bold py-2 px-4 rounded-lg hover:bg-gray-300 focus:outline-none focus:ring-2 focus:ring-gray-400 focus:ring-opacity-50 transition-colors duration-200"
                >
                    Example
                </button>
                <div className="relative w-1/2 group">
                    <button
                        onClick={() => setShowConfirmModal(true)}
                        disabled={!isFormValid()}
                        className="w-full bg-blue-600 text-white font-bold py-2 px-4 rounded-lg hover:bg-blue-700 focus:outline-none focus:ring-2 focus:ring-blue-500 focus:ring-opacity-50 transition-colors duration-200 disabled:bg-gray-400 disabled:cursor-not-allowed"
                    >
                        Send Email
                    </button>
                    {!isFormValid() && (
                        <span className="hidden group-hover:block absolute bottom-full left-1/2 transform -translate-x-1/2 mb-2 w-56 bg-gray-800 text-white text-xs rounded-md p-2 text-center">
                            {getValidationMessage()}
                        </span>
                    )}
                </div>
            </div>
            <p className="text-xs text-gray-500 mt-2 h-4 text-center">{status}</p>

            {/* Modals */}
            <ExampleModal
                isOpen={showExampleModal}
                onClose={() => setShowExampleModal(false)}
                studentData={studentDataCache}
                fromTemplate={fromPills[0] || ''}
                ccRecipients={ccPills}
                subjectTemplate={subject}
                bodyTemplate={body}
            />

            <TemplatesModal
                isOpen={showTemplatesModal}
                onClose={() => setShowTemplatesModal(false)}
                onLoadTemplate={(template) => {
                    setFromPills([template.from]);
                    setSubject(template.subject);
                    setBody(template.body);
                    setCcPills(template.cc || []);
                }}
            />

            <CustomParamModal
                isOpen={showCustomParamModal}
                onClose={() => setShowCustomParamModal(false)}
                customParameters={customParameters}
                onSave={saveCustomParameters}
            />

            <RecipientModal
                isOpen={showRecipientModal}
                onClose={() => setShowRecipientModal(false)}
                currentSelection={recipientSelection}
                onConfirm={handleRecipientUpdate}
                getStudentDataCore={getStudentDataCore}
                recipientDataCache={recipientDataCache}
                onDataFetch={getStudentDataWithUI}
            />

            <ConfirmSendModal
                isOpen={showConfirmModal}
                onClose={() => setShowConfirmModal(false)}
                onConfirm={executeSend}
                count={studentDataCache.length}
            />

            <SuccessModal
                isOpen={showSuccessModal}
                onClose={() => setShowSuccessModal(false)}
                count={lastSentPayload.length}
                payload={lastSentPayload}
                bodyTemplate={body}
            />
        </div>
    );
}
