import React, { useState, useEffect, useRef, useCallback } from 'react';
import ReactQuill from 'react-quill-new';
import 'react-quill-new/dist/quill.snow.css';
import PillInput from './components/PillInput';
import { EMAIL_TEMPLATES_KEY, CUSTOM_PARAMS_KEY, standardParameters, specialParameters, QUILL_EDITOR_CONFIG, PARAMETER_BUTTON_STYLES, COLUMN_MAPPINGS } from './utils/constants';
import { findColumnIndex, getTodaysLdaSheetName, getNameParts, isValidEmail, isValidHttpUrl, evaluateMapping, renderTemplate, renderCCTemplate, generateMissingAssignmentsList } from './utils/helpers';
import { generatePdfReceipt } from './utils/receiptGenerator';
import ExampleModal from './modals/ExampleModal';
import TemplatesModal from './modals/TemplatesModal';
import CustomParamModal from './modals/CustomParamModal';
import RecipientModal from './modals/RecipientModal';
import ConfirmSendModal from './modals/ConfirmSendModal';
import SuccessModal from './modals/SuccessModal';

export default function PersonalizedEmail({ user, accessToken, onReady }) {
    // Connection state
    const [powerAutomateConnection, setPowerAutomateConnection] = useState(null);
    const [isConnected, setIsConnected] = useState(false);
    const [setupUrl, setSetupUrl] = useState('');
    const [setupStatus, setSetupStatus] = useState('');
    const [mode, setMode] = useState(null); // null (loading), 'individual', or 'powerautomate'
    const [userEmail, setUserEmail] = useState('');
    const [localAccessToken, setLocalAccessToken] = useState(accessToken);
    const [consentStatus, setConsentStatus] = useState(null); // null, 'checking', 'granted', 'required'

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
    const [recipientCount, setRecipientCount] = useState(0);
    const [recipientDataCache, setRecipientDataCache] = useState(new Map());
    const [worksheetDataCache, setWorksheetDataCache] = useState({});

    // UI state
    const [lastFocusedInput, setLastFocusedInput] = useState(null);
    const [showMoreParams, setShowMoreParams] = useState(false);
    const [showRecipientHighlight, setShowRecipientHighlight] = useState(false);
    const [showSendContextMenu, setShowSendContextMenu] = useState(false);
    const quillRef = useRef(null);
    const recipientButtonRef = useRef(null);
    const sendButtonRef = useRef(null);

    // Modal states
    const [showExampleModal, setShowExampleModal] = useState(false);
    const [showTemplatesModal, setShowTemplatesModal] = useState(false);
    const [showCustomParamModal, setShowCustomParamModal] = useState(false);
    const [showRecipientModal, setShowRecipientModal] = useState(false);
    const [showConfirmModal, setShowConfirmModal] = useState(false);
    const [showSuccessModal, setShowSuccessModal] = useState(false);
    const [lastSentPayload, setLastSentPayload] = useState([]);

    // Pre-loaded templates state
    const [templates, setTemplates] = useState(null); // null = loading, [] = loaded (empty or with data)

    // Check for existing connection on mount
    useEffect(() => {
        const initializeComponent = async () => {
            // Load user email from localStorage
            try {
                const userInfoStr = localStorage.getItem('SSO_USER_INFO');
                if (userInfoStr) {
                    const userInfo = JSON.parse(userInfoStr);
                    setUserEmail((userInfo.email || '').toLowerCase());
                }
            } catch (error) {
                console.error('Error loading user email:', error);
            }

            // Load connection, custom parameters, and templates in parallel
            await Promise.all([
                checkConnection(),
                loadCustomParameters(),
                loadTemplates()
            ]);
            // Call onReady after all initialization is complete
            if (onReady) onReady();
        };
        initializeComponent();
    }, [onReady]);

    // Fetch fresh authentication token on component mount
    useEffect(() => {
        const fetchFreshToken = async () => {
            try {
                if (typeof Office !== 'undefined' && Office.auth) {
                    console.log('Fetching fresh authentication token...');
                    const token = await Office.auth.getAccessToken({
                        allowSignInPrompt: false,
                        forMSGraphAccess: true
                    });
                    setLocalAccessToken(token);
                    console.log('Fresh authentication token obtained');
                } else {
                    console.log('Office.auth not available, using prop token');
                    setLocalAccessToken(accessToken);
                }
            } catch (error) {
                console.error('Failed to get fresh token:', error);
                // Fall back to prop token
                setLocalAccessToken(accessToken);
            }
        };

        fetchFreshToken();
    }, []);

    // Check consent status in individual mode
    useEffect(() => {
        const checkConsentStatus = async () => {
            // Skip if mode hasn't been determined yet
            if (mode === null) {
                return;
            }

            // Only check consent in individual mode
            if (mode !== 'individual') {
                setConsentStatus('granted');
                return;
            }

            // Wait for localAccessToken to be available
            if (!localAccessToken) {
                return;
            }

            setConsentStatus('checking');

            try {
                // Try to exchange token to check if consent is granted
                const tokenExchangeResponse = await fetch('https://student-retention-token-exchange-dnfdg0hxhsa3gjb4.canadacentral-01.azurewebsites.net/api/exchange-token', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ token: localAccessToken })
                });

                const responseData = await tokenExchangeResponse.json();

                // Check if the error is consent-related
                if (!tokenExchangeResponse.ok) {
                    const errorDetails = responseData.error || responseData.details || '';

                    // Check for consent-required errors (AADSTS65001, AADSTS70011, etc.)
                    if (errorDetails.includes('AADSTS65001') ||
                        errorDetails.includes('AADSTS70011') ||
                        errorDetails.includes('consent')) {
                        console.log('Consent required for Graph API access');
                        setConsentStatus('required');
                    } else {
                        // Other errors - assume consent might still be needed
                        console.warn('Token exchange check failed:', errorDetails);
                        setConsentStatus('required');
                    }
                } else {
                    // Token exchange succeeded - consent is granted
                    console.log('Consent already granted');
                    setConsentStatus('granted');
                }
            } catch (error) {
                console.error('Failed to check consent status:', error);
                // On network error, assume we need consent
                setConsentStatus('required');
            }
        };

        // Run check when mode changes or when localAccessToken becomes available
        if (localAccessToken) {
            checkConsentStatus();
        }
    }, [mode, localAccessToken]);

    // Close send context menu when clicking outside
    useEffect(() => {
        const handleClickOutside = (event) => {
            if (showSendContextMenu && sendButtonRef.current && !sendButtonRef.current.parentElement.contains(event.target)) {
                setShowSendContextMenu(false);
            }
        };

        document.addEventListener('mousedown', handleClickOutside);
        return () => document.removeEventListener('mousedown', handleClickOutside);
    }, [showSendContextMenu]);

    // Setup automatic parameter highlighting
    useEffect(() => {
        if (quillRef.current) {
            const editor = quillRef.current.getEditor();

            const handleTextChange = () => {
                const text = editor.getText();
                const paramRegex = /\{([a-zA-Z0-9_]+)\}/g;
                let match;
                const paramPositions = [];

                // Find all parameter positions
                while ((match = paramRegex.exec(text)) !== null) {
                    paramPositions.push({
                        start: match.index,
                        length: match[0].length,
                        name: match[1]
                    });
                }

                // Apply formatting to each parameter
                paramPositions.forEach(param => {
                    const colors = getParameterColor(param.name);
                    const currentFormat = editor.getFormat(param.start, param.length);

                    // Only update if background or text color is different
                    if (currentFormat.background !== colors.background || currentFormat.color !== colors.color) {
                        editor.formatText(param.start, param.length, {
                            background: colors.background,
                            color: colors.color
                        }, 'silent');
                    }
                });

                // Remove formatting from incomplete parameters (e.g., just typed '{' or removed '}')
                const allText = editor.getText();
                const contents = editor.getContents();
                let index = 0;

                contents.ops.forEach(op => {
                    if (op.insert && typeof op.insert === 'string') {
                        const opText = op.insert;

                        // Check if this text has a background but doesn't contain a complete parameter
                        if (op.attributes && (op.attributes.background || op.attributes.color)) {
                            const hasCompleteParam = /\{[a-zA-Z0-9_]+\}/.test(opText);
                            if (!hasCompleteParam) {
                                // Remove background and color from incomplete parameter text
                                editor.formatText(index, opText.length, {
                                    background: false,
                                    color: false
                                }, 'silent');
                            }
                        }
                        index += opText.length;
                    }
                });
            };

            editor.on('text-change', handleTextChange);

            return () => {
                editor.off('text-change', handleTextChange);
            };
        }
    }, [customParameters]); // Re-run when custom parameters change

    // Pre-cache recipient data after connection is established
    useEffect(() => {
        if (isConnected) {
            preCacheRecipientData();
        }
    }, [isConnected]);

    // Auto-populate From field in individual mode
    useEffect(() => {
        if (mode === 'individual' && userEmail && fromPills.length === 0) {
            setFromPills([userEmail]);
        }
    }, [mode, userEmail, fromPills.length]);

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
                setMode('powerautomate');
            } else {
                setIsConnected(false);
                setMode('individual');
                // In individual mode, auto-populate From field with user's email
                if (userEmail && fromPills.length === 0) {
                    setFromPills([userEmail]);
                }
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

    const loadTemplates = async () => {
        try {
            const loadedTemplates = await Excel.run(async (context) => {
                const settings = context.workbook.settings;
                const templatesSetting = settings.getItemOrNullObject(EMAIL_TEMPLATES_KEY);
                templatesSetting.load("value");
                await context.sync();
                return templatesSetting.value ? JSON.parse(templatesSetting.value) : [];
            });
            setTemplates(loadedTemplates);
        } catch (error) {
            console.error('Error loading templates:', error);
            setTemplates([]); // Set to empty array on error so we don't block loading
        }
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

    const getStudentDataCore = async (selection, skipSpecialParams = false, specialParamsToProcess = []) => {
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
                const usedRange = sheet.getUsedRangeOrNullObject();
                usedRange.load("isNullObject");
                await context.sync();

                if (usedRange.isNullObject) {
                    const err = new Error(`Sheet "${sheetName}" is empty — no data found.`);
                    err.userFacing = true;
                    throw err;
                }

                // Only load cell properties (fill colors) when the exclusion toggle is on
                let cellProperties = null;
                if (excludeFillColor) {
                    cellProperties = usedRange.getCellProperties({ format: { fill: { color: true } } });
                }

                // Only load formulas if we need to process special parameters
                if (skipSpecialParams) {
                    usedRange.load("values");
                } else {
                    usedRange.load("values, formulas");
                }
                await context.sync();

                const values = usedRange.values;
                const formulas = skipSpecialParams ? null : usedRange.formulas;
                const formats = cellProperties ? cellProperties.value : null;
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

                    if (excludeFillColor && formats && colIndices.Outreach !== -1) {
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
                        const colIndex = customParamIndices[param.name];
                        if (colIndex !== undefined) {
                            const cellValue = row[colIndex] ?? '';
                            // If no mappings configured, use the raw cell value
                            if (!param.mappings || param.mappings.length === 0) {
                                value = cellValue;
                            } else if (cellValue !== '') {
                                // If mappings exist, try to find a match
                                for (const mapping of param.mappings) {
                                    if (evaluateMapping(cellValue, mapping)) {
                                        value = mapping.then;
                                        break;
                                    }
                                }
                                // If no mapping matches, leave value blank (empty string)
                            }
                        }
                        student[param.name] = value;
                    }

                    // Process special parameters (only when generating actual emails, not during counting)
                    if (!skipSpecialParams && specialParamsToProcess.length > 0) {
                        for (const paramName of specialParamsToProcess) {
                            if (paramName === 'MissingAssignmentsList') {
                                const gradeBookIndex = colIndices.GradeBook;
                                if (gradeBookIndex !== -1) {
                                    const gradeBookValue = row[gradeBookIndex];
                                    const gradeBookFormula = formulas[i][gradeBookIndex];
                                    student.MissingAssignmentsList = await generateMissingAssignmentsList(
                                        gradeBookValue,
                                        gradeBookFormula,
                                        context
                                    );
                                } else {
                                    student.MissingAssignmentsList = '';
                                }
                            } else if (paramName === 'DaysLeft') {
                                // Calculate DaysLeft as 14 - DaysOut, defaulting to 0 if negative
                                const daysOut = parseInt(student.DaysOut, 10) || 0;
                                const daysLeft = Math.max(0, 14 - daysOut);
                                student.DaysLeft = daysLeft.toString();
                            } else if (paramName === 'Salutation') {
                                // Get time-based greeting
                                const hour = new Date().getHours();
                                let timeGreeting;
                                if (hour < 12) {
                                    timeGreeting = 'Good Morning';
                                } else if (hour < 17) {
                                    timeGreeting = 'Good Afternoon';
                                } else {
                                    timeGreeting = 'Good Evening';
                                }
                                // Randomly pick from greetings including time-based one
                                const greetings = ['Dear', 'Hello', 'Greetings', timeGreeting];
                                student.Salutation = greetings[Math.floor(Math.random() * greetings.length)];
                            }
                        }
                    }

                    includedStudents.push(student);
                }
            });
            return { included: includedStudents, excluded: excludedStudents };
        } catch (error) {
            if (error.code === 'ItemNotFound') {
                error.userFacingMessage = `Error: Sheet "${sheetName}" not found.`;
            } else if (error.userFacing) {
                error.userFacingMessage = error.message;
            } else if (error.code === 'InvalidArgument' || error.code === 'GeneralException') {
                error.userFacingMessage = `Error reading "${sheetName}": ${error.message || 'The sheet may be too large or empty.'}`;
            } else {
                error.userFacingMessage = `Error: ${error.message || 'An unknown error occurred while reading the sheet.'}`;
            }
            console.error(`[PersonalizedEmail] getStudentDataCore error on sheet "${sheetName}":`, error);
            throw error;
        }
    };

    const getStudentDataWithUI = async () => {
        try {
            // Determine which special parameters are actually used in the email template
            const specialParamsToProcess = specialParameters.filter(param =>
                isParameterUsedInTemplate(param)
            );

            // Set specific status message based on what's being fetched
            if (specialParamsToProcess.includes('MissingAssignmentsList')) {
                setStatus('Fetching students and missing assignments...');
            } else {
                setStatus('Fetching students...');
            }

            const result = await getStudentDataCore(recipientSelection, false, specialParamsToProcess);
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
            const ldaResult = await getStudentDataCore(ldaSelection, true); // Skip special params for faster caching
            setRecipientDataCache(prev => new Map(prev).set('lda', ldaResult));

            const masterSelection = { type: 'master', customSheetName: '', excludeDNC: true, excludeFillColor: true };
            const masterResult = await getStudentDataCore(masterSelection, true); // Skip special params for faster caching
            setRecipientDataCache(prev => new Map(prev).set('master', masterResult));
        } catch (error) {
            console.warn("Pre-caching failed. This may happen if sheets are not yet created.", error);
        }
    };

    const handleRecipientUpdate = (newSelection, count) => {
        setRecipientSelection({ ...newSelection, hasBeenSet: true });
        setRecipientCount(count);
        // Clear cache to ensure fresh data is fetched when needed
        setStudentDataCache([]);
    };

    const isParameterUsedInTemplate = (paramName) => {
        // Check if the parameter is used in any part of the email template
        const template = `${fromPills[0] || ''} ${ccPills.join(' ')} ${subject} ${body}`;
        return template.includes(`{${paramName}}`);
    };

    const ensureStudentDataLoaded = async () => {
        // Only fetch if we don't have data cached
        if (studentDataCache.length === 0 && recipientSelection.hasBeenSet) {
            await getStudentDataWithUI();
        }
    };

    const handleOpenExampleModal = async () => {
        await ensureStudentDataLoaded();
        setShowExampleModal(true);
    };

    const handleExampleButtonClick = async () => {
        // If no recipients selected, scroll to recipient button and highlight it
        if (recipientCount === 0) {
            recipientButtonRef.current?.scrollIntoView({ behavior: 'smooth', block: 'center' });
            setShowRecipientHighlight(true);
            setTimeout(() => setShowRecipientHighlight(false), 2000);
            return;
        }
        // Otherwise, open the example modal
        await handleOpenExampleModal();
    };

    const handleOpenConfirmModal = async () => {
        await ensureStudentDataLoaded();
        setShowConfirmModal(true);
    };

    const getParameterColor = (paramName) => {
        // Check if it's a special parameter
        if (specialParameters.includes(paramName)) {
            return {
                background: '#fed7aa', // orange-200
                color: '#9a3412'       // orange-800
            };
        }

        // Check if it's a custom parameter
        const customParam = customParameters.find(p => p.name === paramName);
        if (customParam) {
            const hasMappings = customParam.mappings && customParam.mappings.length > 0;
            const hasNested = hasMappings && customParam.mappings.some(m => /\{(\w+)\}/.test(m.then));

            if (hasNested) {
                return {
                    background: '#fecdd3', // rose-200
                    color: '#881337'       // rose-800
                };
            }
            if (hasMappings) {
                return {
                    background: '#e9d5ff', // purple-200
                    color: '#581c87'       // purple-800
                };
            }
            return {
                background: '#bfdbfe', // blue-200
                color: '#1e3a8a'       // blue-800
            };
        }

        // Standard parameter
        return {
            background: '#e5e7eb', // gray-200
            color: '#374151'       // gray-700
        };
    };

    const insertParameter = (param) => {
        if (lastFocusedInput === 'quill' && quillRef.current) {
            const editor = quillRef.current.getEditor();
            const range = editor.getSelection(true);

            // Extract parameter name from {ParamName} format
            const paramName = param.replace(/[{}]/g, '');
            const colors = getParameterColor(paramName);

            // Insert with background and text color formatting
            editor.insertText(range.index, param, {
                background: colors.background,
                color: colors.color
            });

            // Insert a zero-width space after the parameter without formatting
            // This creates a buffer that prevents the cursor from picking up the highlight
            const bufferPosition = range.index + param.length;
            editor.insertText(bufferPosition, '\u200B', {
                background: false,
                color: false
            });

            // Move cursor after the buffer
            editor.setSelection(bufferPosition + 1);
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

            const paramName = param.replace(/[{}]/g, '');
            const colors = getParameterColor(paramName);

            const insertPosition = editor.getLength() - 1;
            editor.insertText(insertPosition, param, {
                background: colors.background,
                color: colors.color
            });

            // Insert a zero-width space after the parameter without formatting
            const bufferPosition = insertPosition + param.length;
            editor.insertText(bufferPosition, '\u200B', {
                background: false,
                color: false
            });

            // Move cursor after the buffer
            editor.setSelection(bufferPosition + 1);
        }
    };

    const stripParameterBackgrounds = (html) => {
        // Create a temporary div to parse HTML
        const tempDiv = document.createElement('div');
        tempDiv.innerHTML = html;

        // Find all elements with background or color styling
        const styledElements = tempDiv.querySelectorAll('[style*="background"], [style*="color"]');

        styledElements.forEach(element => {
            const text = element.textContent || '';
            // Only strip background and color if this element contains a parameter pattern
            if (/\{[a-zA-Z0-9_]+\}/.test(text)) {
                // Remove background-color and color from inline style
                const style = element.getAttribute('style') || '';
                const newStyle = style
                    .replace(/background-color:\s*[^;]+;?/gi, '')
                    .replace(/background:\s*[^;]+;?/gi, '')
                    .replace(/color:\s*[^;]+;?/gi, '')
                    .trim();

                if (newStyle) {
                    element.setAttribute('style', newStyle);
                } else {
                    element.removeAttribute('style');
                }
            }
        });

        return tempDiv.innerHTML;
    };

    const generatePayload = () => {
        const fromTemplate = fromPills[0] || '';
        // Strip parameter backgrounds from body before rendering
        const cleanBodyHtml = stripParameterBackgrounds(body);

        const emails = studentDataCache.map(student => ({
            from: renderTemplate(fromTemplate, student),
            to: student.StudentEmail || '',
            cc: renderCCTemplate(ccPills, student),
            subject: renderTemplate(subject, student),
            body: renderTemplate(cleanBodyHtml, student)
        })).filter(email => email.to && email.from);

        // Calculate sender breakdown
        const senderCounts = emails.reduce((acc, email) => {
            const from = email.from || 'Unknown';
            acc[from] = (acc[from] || 0) + 1;
            return acc;
        }, {});

        const senders = Object.entries(senderCounts).map(([email, count]) => ({
            email,
            count
        }));

        return {
            byName: user || '',
            byEmail: userEmail || '',
            totalCount: emails.length,
            senders: senders,
            emails: emails
        };
    };

    const showConsentDialog = () => {
        return new Promise((resolve, reject) => {
            const dialogUrl = 'https://vsblanco.github.io/Student-Retention-Add-in/consent-dialog.html';

            Office.context.ui.displayDialogAsync(
                dialogUrl,
                { height: 60, width: 30, promptBeforeOpen: false },
                (result) => {
                    if (result.status === Office.AsyncResultStatus.Failed) {
                        reject(new Error('Failed to open consent dialog'));
                        return;
                    }

                    const dialog = result.value;

                    dialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg) => {
                        dialog.close();

                        try {
                            const response = JSON.parse(arg.message);

                            if (response.status === 'success') {
                                // User consented successfully
                                resolve({ success: true });
                            } else {
                                reject(new Error(response.description || response.error || 'Consent failed'));
                            }
                        } catch (e) {
                            reject(new Error('Invalid response from consent dialog'));
                        }
                    });

                    dialog.addEventHandler(Office.EventType.DialogEventReceived, (arg) => {
                        dialog.close();
                        reject(new Error('Dialog was closed'));
                    });
                }
            );
        });
    };

    const sendEmailsViaGraphAPI = async () => {
        setShowConfirmModal(false);

        const payload = generatePayload();
        setLastSentPayload(payload);

        if (payload.emails.length === 0) {
            setStatus('No students with valid "To" and "From" email addresses found.');
            return;
        }

        if (!localAccessToken) {
            setStatus('Authentication token not available. Please log in again.');
            return;
        }

        let successCount = 0;
        let failureCount = 0;
        const errors = [];

        try {
            // Step 1: Get a fresh SSO token
            setStatus('Getting authentication token...');
            const newToken = await Office.auth.getAccessToken({
                allowSignInPrompt: false,
                forMSGraphAccess: true
            });

            // Step 2: Exchange Office SSO token for Graph API token
            setStatus('Exchanging authentication token...');
            const tokenExchangeResponse = await fetch('https://student-retention-token-exchange-dnfdg0hxhsa3gjb4.canadacentral-01.azurewebsites.net/api/exchange-token', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ token: newToken })
            });

            const responseData = await tokenExchangeResponse.json();

            if (!tokenExchangeResponse.ok) {
                throw new Error(`Token exchange failed: ${responseData.error || responseData.details || 'Unknown error'}`);
            }

            const { accessToken: graphToken } = responseData;

            // Step 3: Send emails using Graph API token
            await sendEmailsWithGraphToken(graphToken, payload.emails, setStatus, successCount, failureCount, errors);

        } catch (error) {
            setStatus(`Failed to send emails: ${error.message}`);
            console.error("Error sending emails:", error);
        }
    };

    const sendEmailsWithGraphToken = async (graphToken, emails, setStatus, successCount, failureCount, errors) => {
        setStatus(`Sending ${emails.length} emails...`);

        for (const email of emails) {
            try {
                // Parse CC recipients
                const ccRecipients = email.cc
                    ? email.cc.split(',').map(addr => addr.trim()).filter(addr => addr).map(addr => ({
                        emailAddress: { address: addr }
                    }))
                    : [];

                // Construct Microsoft Graph API sendMail payload
                const graphPayload = {
                    message: {
                        subject: email.subject,
                        body: {
                            contentType: 'HTML',
                            content: email.body
                        },
                        toRecipients: [
                            {
                                emailAddress: {
                                    address: email.to
                                }
                            }
                        ],
                        ccRecipients: ccRecipients
                    },
                    saveToSentItems: true
                };

                const response = await fetch('https://graph.microsoft.com/v1.0/me/sendMail', {
                    method: 'POST',
                    headers: {
                        'Authorization': `Bearer ${graphToken}`,
                        'Content-Type': 'application/json'
                    },
                    body: JSON.stringify(graphPayload)
                });

                if (!response.ok) {
                    const errorText = await response.text();
                    throw new Error(`HTTP ${response.status}: ${errorText}`);
                }

                successCount++;
                setStatus(`Sent ${successCount} of ${emails.length} emails...`);
            } catch (error) {
                failureCount++;
                errors.push({ to: email.to, error: error.message });
                console.error(`Failed to send email to ${email.to}:`, error);
            }
        }

        if (failureCount === 0) {
            setStatus(`Successfully sent ${successCount} emails!`);
            setShowSuccessModal(true);
        } else {
            setStatus(`Sent ${successCount} emails. Failed: ${failureCount}. Check console for details.`);
            console.error('Email sending errors:', errors);
        }
    };

    const sendEmailsViaPowerAutomate = async () => {
        setShowConfirmModal(false);
        setStatus(`Sending ${studentDataCache.length} emails...`);

        const payload = generatePayload();

        if (payload.emails.length === 0) {
            setStatus('No students with valid "To" and "From" email addresses found.');
            return;
        }

        // Generate base64 PDF receipt to include in payload
        const initiator = { name: user, email: userEmail };
        const receiptBase64 = generatePdfReceipt(
            payload.emails,
            body,
            initiator,
            true // returnBase64
        );

        // Add receipt to payload
        const payloadWithReceipt = {
            ...payload,
            receipt: receiptBase64 || ''
        };

        setLastSentPayload(payloadWithReceipt);

        try {
            const response = await fetch(powerAutomateConnection.url, {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify(payloadWithReceipt)
            });
            if (!response.ok) throw new Error(`HTTP error! status: ${response.status}`);
            setStatus(`Successfully sent ${payload.emails.length} emails!`);
            setShowSuccessModal(true);
        } catch (error) {
            setStatus(`Failed to send emails: ${error.message}`);
            console.error("Error sending emails:", error);
        }
    };

    const executeSend = async () => {
        if (mode === 'individual') {
            await sendEmailsViaGraphAPI();
        } else if (mode === 'powerautomate') {
            await sendEmailsViaPowerAutomate();
        } else {
            setStatus('Invalid sending mode. Please refresh and try again.');
        }
    };

    const handleSendTestEmail = async () => {
        setShowSendContextMenu(false);

        // Validate we have minimum required fields (From, Subject, Body)
        const from = fromPills[0] || '';
        if (!from.trim() || !subject.trim() || !body.trim()) {
            setStatus('Please fill in From, Subject, and Body to send a test email.');
            return;
        }

        if (!userEmail) {
            setStatus('Unable to determine your email address. Please sign in again.');
            return;
        }

        // Ensure student data is loaded
        await ensureStudentDataLoaded();

        // Pick a random student for parameter replacement, or use a placeholder if no students
        let testStudent;
        if (studentDataCache.length > 0) {
            const randomIndex = Math.floor(Math.random() * studentDataCache.length);
            testStudent = studentDataCache[randomIndex];
        } else {
            // Create placeholder student data if no students selected
            testStudent = {
                StudentEmail: userEmail,
                FirstName: 'Test',
                LastName: 'Student',
                StudentName: 'Test Student'
            };
        }

        // Generate single test email payload with user's email as recipient
        const fromTemplate = fromPills[0] || '';
        const cleanBodyHtml = stripParameterBackgrounds(body);
        const testFromEmail = renderTemplate(fromTemplate, testStudent);

        const testPayload = {
            byName: user || '',
            byEmail: userEmail || '',
            totalCount: 1,
            senders: [{ email: testFromEmail, count: 1 }],
            emails: [{
                from: testFromEmail,
                to: userEmail, // Always send to the logged-in user
                cc: '', // Don't CC anyone on test emails
                subject: `[TEST] ${renderTemplate(subject, testStudent)}`,
                body: renderTemplate(cleanBodyHtml, testStudent)
            }]
        };

        setLastSentPayload(testPayload);
        setStatus('Sending test email to yourself...');

        try {
            if (mode === 'individual') {
                // Get fresh token and send via Graph API
                const newToken = await Office.auth.getAccessToken({
                    allowSignInPrompt: false,
                    forMSGraphAccess: true
                });

                const tokenExchangeResponse = await fetch('https://student-retention-token-exchange-dnfdg0hxhsa3gjb4.canadacentral-01.azurewebsites.net/api/exchange-token', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ token: newToken })
                });

                const responseData = await tokenExchangeResponse.json();
                if (!tokenExchangeResponse.ok) {
                    throw new Error(`Token exchange failed: ${responseData.error || responseData.details || 'Unknown error'}`);
                }

                const { accessToken: graphToken } = responseData;

                // Send single test email
                const email = testPayload.emails[0];
                const ccRecipients = email.cc
                    ? email.cc.split(',').map(addr => addr.trim()).filter(addr => addr).map(addr => ({
                        emailAddress: { address: addr }
                    }))
                    : [];

                const graphPayload = {
                    message: {
                        subject: email.subject,
                        body: { contentType: 'HTML', content: email.body },
                        toRecipients: [{ emailAddress: { address: email.to } }],
                        ccRecipients: ccRecipients
                    },
                    saveToSentItems: true
                };

                const response = await fetch('https://graph.microsoft.com/v1.0/me/sendMail', {
                    method: 'POST',
                    headers: {
                        'Authorization': `Bearer ${graphToken}`,
                        'Content-Type': 'application/json'
                    },
                    body: JSON.stringify(graphPayload)
                });

                if (!response.ok) {
                    const errorText = await response.text();
                    throw new Error(`HTTP ${response.status}: ${errorText}`);
                }

                setStatus(`Test email sent to ${userEmail}!`);
            } else if (mode === 'powerautomate') {
                // Generate base64 PDF receipt for test email
                const initiator = { name: user, email: userEmail };
                const receiptBase64 = generatePdfReceipt(
                    testPayload.emails,
                    body,
                    initiator,
                    true // returnBase64
                );

                // Add receipt to payload
                const testPayloadWithReceipt = {
                    ...testPayload,
                    receipt: receiptBase64 || ''
                };

                // Send via Power Automate
                const response = await fetch(powerAutomateConnection.url, {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify(testPayloadWithReceipt)
                });

                if (!response.ok) throw new Error(`HTTP error! status: ${response.status}`);
                setStatus(`Test email sent to ${userEmail}!`);
            } else {
                setStatus('Invalid sending mode. Please refresh and try again.');
            }
        } catch (error) {
            setStatus(`Failed to send test email: ${error.message}`);
            console.error("Error sending test email:", error);
        }
    };

    const isFormValid = () => {
        const from = fromPills[0] || '';
        const isFromValid = from && from.trim() !== '';
        const isSubjectValid = subject && subject.trim() !== '';
        const isBodyValid = body && body.trim() !== '';
        const areRecipientsValid = recipientSelection.hasBeenSet && recipientCount > 0;
        return isFromValid && isSubjectValid && isBodyValid && areRecipientsValid;
    };

    const getValidationMessage = () => {
        const missing = [];
        if (!fromPills[0] || !fromPills[0].trim()) missing.push('From address');
        if (!recipientSelection.hasBeenSet || recipientCount === 0) missing.push('Recipients');
        if (!subject || !subject.trim()) missing.push('Subject');
        if (!body || !body.trim()) missing.push('Body');
        return missing.length > 0 ? `Required: ${missing.join(', ')}.` : '';
    };

    const renderParameterButton = (param) => {
        const isCustom = typeof param === 'object';
        const paramName = isCustom ? param.name : param;

        let buttonClass = PARAMETER_BUTTON_STYLES.standard;

        // Check if this is a special parameter
        if (specialParameters.includes(paramName)) {
            buttonClass = PARAMETER_BUTTON_STYLES.special;
        } else if (isCustom) {
            const hasMappings = param.mappings && param.mappings.length > 0;
            const hasNested = hasMappings && param.mappings.some(m => /\{(\w+)\}/.test(m.then));

            if (hasNested) buttonClass = PARAMETER_BUTTON_STYLES.nested;
            else if (hasMappings) buttonClass = PARAMETER_BUTTON_STYLES.mapped;
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

    // Show loading screen while checking connection/mode or loading templates
    if (mode === null || templates === null) {
        return (
            <div className="max-w-md mx-auto p-4 bg-gray-50 min-h-screen flex items-center justify-center">
                <div className="text-center">
                    <div className="animate-spin rounded-full h-12 w-12 border-b-2 border-blue-600 mx-auto mb-4"></div>
                    <p className="text-gray-600">Loading...</p>
                </div>
            </div>
        );
    }

    // Show loading screen while checking consent
    if (consentStatus === 'checking') {
        return (
            <div className="max-w-md mx-auto p-4 bg-gray-50 min-h-screen flex items-center justify-center">
                <div className="text-center">
                    <div className="animate-spin rounded-full h-12 w-12 border-b-2 border-blue-600 mx-auto mb-4"></div>
                    <p className="text-gray-600">Checking permissions...</p>
                </div>
            </div>
        );
    }

    // Show consent required screen if admin consent is needed
    if (consentStatus === 'required') {
        const adminConsentUrl = 'https://login.microsoftonline.com/common/adminconsent?client_id=71f37f39-a330-413a-be61-0baa5ce03ea3';

        return (
            <div className="max-w-md mx-auto p-4 bg-gray-50 min-h-screen flex items-center justify-center">
                <div className="bg-white rounded-lg shadow-lg p-8 max-w-lg">
                    <div className="text-center mb-6">
                        <div className="mx-auto flex items-center justify-center h-16 w-16 rounded-full bg-yellow-100 mb-4">
                            <svg className="h-8 w-8 text-yellow-600" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                                <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M12 9v2m0 4h.01m-6.938 4h13.856c1.54 0 2.502-1.667 1.732-3L13.732 4c-.77-1.333-2.694-1.333-3.464 0L3.34 16c-.77 1.333.192 3 1.732 3z" />
                            </svg>
                        </div>
                        <h2 className="text-2xl font-bold text-gray-900 mb-2">Admin Approval Required</h2>
                        <p className="text-gray-600 mb-6">
                            This feature requires an IT Administrator to enable it before you can send personalized emails using Microsoft Graph API.
                        </p>
                    </div>

                    <div className="bg-blue-50 border border-blue-200 rounded-md p-4 mb-6">
                        <p className="text-sm text-blue-800 mb-2">
                            <strong>For IT Administrators:</strong>
                        </p>
                        <p className="text-sm text-blue-700 mb-3">
                            Click the button below to grant consent for this application. This is a one-time setup that will enable all users in your organization to use this feature.
                        </p>
                        <a
                            href={adminConsentUrl}
                            target="_blank"
                            rel="noopener noreferrer"
                            className="inline-block w-full text-center px-4 py-2 bg-blue-600 text-white font-semibold rounded-md hover:bg-blue-700 transition-colors"
                        >
                            Grant Admin Consent
                        </a>
                    </div>

                    <div className="border-t border-gray-200 pt-4">
                        <p className="text-xs text-gray-500 text-center mb-3">
                            After consent is granted, refresh this page to start sending emails.
                        </p>
                        <p className="text-xs text-gray-500 text-center">
                            <strong>Alternative:</strong> You can configure Power Automate in Settings to send emails without requiring admin approval.
                        </p>
                    </div>
                </div>
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
                    <label className="block text-sm font-medium text-gray-700">
                        From {mode === 'individual' && <span className="text-xs text-gray-500">(Individual Mode)</span>}
                    </label>
                    <PillInput
                        pills={fromPills}
                        onPillsChange={setFromPills}
                        placeholder="sender@example.com"
                        singleValue={true}
                        onFocus={() => setLastFocusedInput('from')}
                        readOnly={mode === 'individual'}
                        noWrap={true}
                    />
                </div>

                {/* Recipient Selection */}
                <div>
                    <label className="block text-sm font-medium text-gray-700">Recipient</label>
                    <button
                        ref={recipientButtonRef}
                        onClick={() => setShowRecipientModal(true)}
                        className={`mt-1 w-full text-left px-3 py-2 border border-gray-300 rounded-md shadow-sm focus:outline-none focus:ring-2 focus:ring-blue-500 transition-all duration-500 ${
                            recipientSelection.hasBeenSet && recipientCount > 0
                                ? 'bg-green-100 text-green-800 font-semibold'
                                : 'bg-white'
                        } ${
                            showRecipientHighlight ? 'ring-2 ring-red-300 bg-red-50' : ''
                        }`}
                    >
                        {recipientSelection.hasBeenSet && recipientCount > 0
                            ? `${recipientCount} Student${recipientCount !== 1 ? 's' : ''} Selected`
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
                        noWrap={true}
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
                        style={{ height: '192px', marginBottom: '80px' }}
                    />
                </div>

                {/* Parameters */}
                <div className="mt-8">
                    <label className="block text-sm font-medium text-gray-700">Insert Parameter</label>
                    <div className="mt-2 flex flex-wrap gap-2">
                        {standardParameters.map(param => renderParameterButton(param))}
                    </div>

                    {specialParameters.length > 0 && (
                        <div className="mt-3">
                            <label className="block text-xs font-medium text-gray-600 mb-2">Special Parameters</label>
                            <div className="flex flex-wrap gap-2">
                                {specialParameters.map(param => renderParameterButton(param))}
                            </div>
                        </div>
                    )}

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
                    onClick={handleExampleButtonClick}
                    disabled={recipientCount === 0}
                    className={`${user === 'Guest' ? 'w-full' : 'w-1/2'} font-bold py-2 px-4 rounded-lg focus:outline-none focus:ring-2 focus:ring-opacity-50 transition-colors duration-200 ${
                        recipientCount === 0
                            ? 'bg-gray-300 text-gray-500 cursor-not-allowed'
                            : 'bg-gray-200 text-gray-800 hover:bg-gray-300 focus:ring-gray-400'
                    }`}
                >
                    Example
                </button>
                {user !== 'Guest' && (
                    <div className="relative w-1/2 group">
                        <button
                            ref={sendButtonRef}
                            onClick={handleOpenConfirmModal}
                            onContextMenu={(e) => {
                                e.preventDefault();
                                setShowSendContextMenu(true);
                            }}
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
                        {/* Context Menu for Send Button */}
                        {showSendContextMenu && (
                            <div className="absolute bottom-full left-0 mb-1 w-full bg-white border border-gray-200 rounded-lg shadow-lg z-50">
                                <button
                                    onClick={handleSendTestEmail}
                                    className="w-full px-4 py-2 text-left text-sm text-gray-700 hover:bg-gray-100 rounded-lg"
                                >
                                    Send Test Email to myself
                                </button>
                            </div>
                        )}
                    </div>
                )}
            </div>
            {mode === 'individual' && user !== 'Guest' && (
                <p className="text-xs text-gray-600 mt-2 text-center">
                    Emails will be sent from your mailbox using Microsoft Graph API
                </p>
            )}
            {mode === 'powerautomate' && user !== 'Guest' && (
                <p className="text-xs text-gray-600 mt-2 text-center">
                    Emails will be sent via Power Automate
                </p>
            )}
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
                user={user}
                userEmail={userEmail}
                currentFrom={fromPills[0] || ''}
                currentSubject={subject}
                currentBody={body}
                currentCC={ccPills}
                templates={templates}
                onTemplatesChange={setTemplates}
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
                count={lastSentPayload?.emails?.length || 0}
                payload={lastSentPayload}
                bodyTemplate={body}
                initiator={{ name: user, email: userEmail }}
            />
        </div>
    );
}
