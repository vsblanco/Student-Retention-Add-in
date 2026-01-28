// V-7.2 - 2025-10-02 - 9:54 AM EDT
import { findColumnIndex, getTodaysLdaSheetName, getNameParts } from './utils.js';
import { EMAIL_TEMPLATES_KEY, CUSTOM_PARAMS_KEY, standardParameters, QUILL_EDITOR_CONFIG, COLUMN_MAPPINGS, PARAMETER_BUTTON_STYLES } from './constants.js';
import ModalManager from './modal.js';
import { generatePdfReceipt } from './receipt-generator.js';

let powerAutomateConnection = null;
let studentDataCache = [];
let lastFocusedInput = null;
let quill;
let fromPill = [];
let ccRecipients = [];
let customParameters = [];
let modalManager;
let worksheetDataCache = {};
let lastSentPayload = [];
let recipientSelection = { type: 'lda', customSheetName: '', excludeDNC: true, excludeFillColor: true, hasBeenSet: false };
let recipientDataCache = new Map(); // Changed from recipientCountCache to store full data objects

/**
 * Validates if a string is a properly formatted email address.
 * @param {string} email The email string to validate.
 * @returns {boolean} True if the email is valid, false otherwise.
 */
function isValidEmail(email) {
    if (typeof email !== 'string' || !email.trim()) {
        return false;
    }
    // A common regex for email validation.
    const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
    return emailRegex.test(email);
}

/**
 * The core data fetching function. It reads data from the specified sheet,
 * processes it, and returns both included and excluded students.
 * @param {object} selection - The recipient selection object.
 * @returns {Promise<{included: Array, excluded: Array}>} A promise that resolves with an object containing included and excluded students.
 */
async function _getStudentDataCore(selection) {
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
    worksheetDataCache = {}; 

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
                                const hasExcludableDnc = individualTags.some(tag => tag.includes('DNC') && !['DNC - PHONE', 'DNC - OTHER PHONE'].includes(tag));

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
            
            await loadCustomParameters();
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

                // ADDED: Exclude students with invalid email addresses first.
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
                    StudentName: studentName, FirstName: nameParts.first, LastName: nameParts.last,
                    StudentEmail: studentEmail, PersonalEmail: row[colIndices.PersonalEmail] ?? '',
                    Grade: row[colIndices.Grade] ?? '', DaysOut: row[colIndices.DaysOut] ?? '', Assigned: row[colIndices.Assigned] ?? ''
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
                            const isAsync = /\bawait\b/.test(userScript), hasReturn = /\breturn\b/.test(userScript);
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
}


Office.onReady((info) => {
    if (info.host === Office.HostType.Excel) {
        quill = new Quill('#editor-container', QUILL_EDITOR_CONFIG);

        async function getStudentDataWithUI() {
            const status = document.getElementById('status');
            status.textContent = 'Fetching students...';
            status.style.color = 'gray';
            try {
                const result = await _getStudentDataCore(recipientSelection);
                studentDataCache = result.included;
                status.textContent = `Found ${studentDataCache.length} students.`;
                validateAllFields();
                setTimeout(() => status.textContent = '', 3000);
                return studentDataCache;
            } catch (error) {
                const message = error.userFacingMessage || (error.userFacing ? error.message : 'An error occurred while fetching data.');
                status.textContent = message;
                status.style.color = 'red';
                validateAllFields();
                throw error;
            }
        }
        
        const appContext = {
            quill,
            setLastFocusedInput: (input) => { lastFocusedInput = input; },
            getStudentDataWithUI,
            getStudentDataCore: _getStudentDataCore,
            updateRecipientSelection: (newSelection, count) => {
                recipientSelection = { ...newSelection, hasBeenSet: true };
                const button = document.getElementById('select-students-button');
                if (count >= 0) {
                    button.textContent = `${count} Student${count !== 1 ? 's' : ''} Selected`;
                    button.classList.add('bg-green-100', 'text-green-800', 'font-semibold');
                } else {
                    button.textContent = 'Select Students';
                    button.classList.remove('bg-green-100', 'text-green-800', 'font-semibold');
                }
                validateAllFields();
            },
            recipientSelection,
            recipientDataCache,
            preCacheRecipientData,
            renderTemplate,
            renderCCTemplate,
            getTemplates,
            saveTemplates,
            loadCustomParameters: async () => {
                customParameters = await getCustomParameters();
                return customParameters;
            },
            getCustomParameters,
            saveCustomParameters,
            populateParameterButtons,
            executeSend,
            ccRecipients,
            renderCCPills,
            get customParameters() { return customParameters; },
            get standardParameters() { return standardParameters; },
            get studentDataCache() { return studentDataCache; }
        };

        modalManager = new ModalManager(appContext);
        
        document.getElementById("send-email-button").onclick = () => modalManager.showSendConfirmModal();
        document.getElementById("create-connection-button").onclick = createConnection;
        document.getElementById("select-students-button").onclick = () => modalManager.showRecipientModal();
        
        setupFromInput();
        setupCcInput();
        setupExampleContextMenu();
        
        const subjectInput = document.getElementById('email-subject');
        subjectInput.addEventListener('focus', () => lastFocusedInput = subjectInput);
        subjectInput.addEventListener('input', validateAllFields);
        
        quill.on('selection-change', (range) => { if (range) lastFocusedInput = quill; });
        quill.on('text-change', validateAllFields);

        loadCustomParameters().then(populateParameterButtons);
        checkConnection().then(preCacheRecipientData);
        validateAllFields();
    }
});

function setupExampleContextMenu() {
    const exampleButton = document.getElementById('show-example-button');
    const contextMenu = document.getElementById('example-context-menu');

    exampleButton.addEventListener('contextmenu', (e) => {
        e.preventDefault();
        const rect = e.target.getBoundingClientRect();
        contextMenu.style.left = `${rect.left}px`;
        contextMenu.style.top = `${rect.bottom}px`;
        contextMenu.classList.remove('hidden');
    });

    document.addEventListener('click', (e) => {
        if (!contextMenu.contains(e.target) && e.target !== exampleButton) {
            contextMenu.classList.add('hidden');
        }
    });
    
    document.getElementById('context-menu-payload').onclick = (e) => {
        e.preventDefault();
        modalManager.showPayloadModal();
        contextMenu.classList.add('hidden');
    };

    document.getElementById('context-menu-receipt').onclick = (e) => {
        e.preventDefault();
        const bodyTemplate = quill.root.innerHTML;
        const payload = generatePayload();
        if (payload.length > 0) {
            generatePdfReceipt(payload, bodyTemplate);
        } else {
            document.getElementById('status').textContent = 'Please select recipients before generating a receipt.';
            document.getElementById('status').style.color = 'orange';
        }
        contextMenu.classList.add('hidden');
    };
}

async function populateParameterButtons() {
    const standardContainer = document.getElementById('standard-parameter-buttons');
    const customContainer = document.getElementById('custom-parameter-buttons');
    const moreCustomContainer = document.getElementById('more-custom-parameters');
    const customSection = document.getElementById('custom-parameter-section');
    const showMoreButton = document.getElementById('show-more-params-button');
    
    standardContainer.innerHTML = ''; 
    customContainer.innerHTML = '';
    moreCustomContainer.innerHTML = '';

    const createButton = (param) => {
        const button = document.createElement('button');
        const isCustom = typeof param === 'object';
        const paramName = isCustom ? param.name : param;
        
        if (isCustom) {
            const hasMappings = param.mappings && param.mappings.length > 0;
            const hasNested = hasMappings && param.mappings.some(m => /\{(\w+)\}/.test(m.then));
            
            if (hasNested) button.className = PARAMETER_BUTTON_STYLES.nested;
            else if (hasMappings) button.className = PARAMETER_BUTTON_STYLES.mapped;
            else button.className = PARAMETER_BUTTON_STYLES.custom;
        } else {
            button.className = PARAMETER_BUTTON_STYLES.standard;
        }

        button.textContent = `{${paramName}}`;
        button.onclick = () => insertParameter(`{${paramName}}`);
        return button;
    };

    standardParameters.forEach(param => {
        standardContainer.appendChild(createButton(param));
    });

    if (customParameters.length > 0) {
        customSection.classList.remove('hidden');
        customParameters.slice(0, 5).forEach(param => customContainer.appendChild(createButton(param)));
        const remainingParams = customParameters.slice(5);

        if (remainingParams.length > 0) {
            remainingParams.forEach(param => moreCustomContainer.appendChild(createButton(param)));
            showMoreButton.textContent = `Show ${remainingParams.length} More...`;
            showMoreButton.classList.remove('hidden');
            showMoreButton.onclick = () => {
                const isHidden = moreCustomContainer.classList.contains('hidden');
                moreCustomContainer.classList.toggle('hidden');
                showMoreButton.textContent = isHidden ? 'Show Less' : `Show ${remainingParams.length} More...`;
            };
        } else {
            showMoreButton.classList.add('hidden');
            moreCustomContainer.classList.add('hidden');
        }
    } else {
        [customSection, showMoreButton, moreCustomContainer].forEach(el => el.classList.add('hidden'));
    }
}

function insertParameter(param) {
    if (lastFocusedInput instanceof Quill) {
        const range = lastFocusedInput.getSelection(true);
        lastFocusedInput.insertText(range.index, param, 'user');
    } else if (lastFocusedInput && lastFocusedInput.id === 'email-from-input') {
        addFromRecipient(param);
    } else if (lastFocusedInput && lastFocusedInput.id === 'email-cc-input') {
        addCcRecipient(param);
    } else if (lastFocusedInput) {
        const start = lastFocusedInput.selectionStart, end = lastFocusedInput.selectionEnd;
        const text = lastFocusedInput.value;
        lastFocusedInput.value = text.substring(0, start) + param + text.substring(end);
        lastFocusedInput.focus();
        lastFocusedInput.selectionStart = lastFocusedInput.selectionEnd = start + param.length;
    } else {
        quill.focus();
        quill.insertText(quill.getLength(), param, 'user');
    }
    validateAllFields();
}

async function checkConnection() {
    await Excel.run(async (context) => {
        const settings = context.workbook.settings;
        const connectionsSetting = settings.getItemOrNullObject("connections");
        connectionsSetting.load("value");
        await context.sync();

        const connections = connectionsSetting.value ? JSON.parse(connectionsSetting.value) : [];
        powerAutomateConnection = connections.find(c => c.type === 'power-automate' && c.name === 'Send Personalized Email');

        if (powerAutomateConnection) {
            document.getElementById('setup-wizard').classList.add('hidden');
            document.getElementById('email-composer').classList.remove('hidden');
        } else {
            document.getElementById('setup-wizard').classList.remove('hidden');
            document.getElementById('email-composer').classList.add('hidden');
        }
    });
}

async function createConnection() {
    const urlInput = document.getElementById('power-automate-url');
    const status = document.getElementById('setup-status');
    const url = urlInput.value.trim();

    if (!isValidHttpUrl(url)) {
        status.textContent = "Please enter a valid HTTP URL.";
        status.style.color = 'red';
        return;
    }
    
    status.textContent = "Creating connection...";
    status.style.color = 'gray';

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
            url: url,
            actions: [],
            history: []
        };
        connections.push(newConnection);
        settings.add("connections", JSON.stringify(connections));
        await context.sync();

        status.textContent = "Connection created successfully!";
        status.style.color = 'green';
        setTimeout(checkConnection, 1500);
    });
}

function evaluateMapping(cellValue, mapping) {
    const cellStr = String(cellValue).trim().toLowerCase();
    const conditionStr = String(mapping.if).trim().toLowerCase();
    const cellNum = parseFloat(cellValue), conditionNum = parseFloat(mapping.if);
    const isNumeric = !isNaN(cellNum) && !isNaN(conditionNum);

    switch (mapping.operator) {
        case 'eq': return cellStr === conditionStr;
        case 'neq': return cellStr !== conditionStr;
        case 'contains': return cellStr.includes(conditionStr);
        case 'does_not_contain': return !cellStr.includes(conditionStr);
        case 'starts_with': return cellStr.startsWith(conditionStr);
        case 'ends_with': return cellStr.endsWith(conditionStr);
        case 'gt': return isNumeric && cellNum > conditionNum;
        case 'lt': return isNumeric && cellNum < conditionNum;
        case 'gte': return isNumeric && cellNum >= conditionNum;
        case 'lte': return isNumeric && cellNum <= conditionNum;
        default: return false;
    }
}

async function getWorksheetData(sheetNameToFetch) {
    if (worksheetDataCache[sheetNameToFetch]) return worksheetDataCache[sheetNameToFetch];
    try {
        await Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getItem(sheetNameToFetch);
            const range = sheet.getUsedRange();
            range.load("values");
            await context.sync();
            const values = range.values;
            worksheetDataCache[sheetNameToFetch] = (values.length > 1)
                ? { headers: values[0].map(h => String(h ?? '').toLowerCase()), values: values.slice(1) }
                : { headers: [], values: [] };
        });
        return worksheetDataCache[sheetNameToFetch];
    } catch (error) {
        if (error.code !== 'ItemNotFound') console.error(`Error loading '${sheetNameToFetch}' sheet:`, error);
        worksheetDataCache[sheetNameToFetch] = null;
        return null;
    }
}

const renderTemplate = (template, data) => {
    if (!template) return '';
    let result = template;
    for (let i = 0; i < 10 && /\{(\w+)\}/.test(result); i++) {
        result = result.replace(/\{(\w+)\}/g, (match, key) => {
            let valueToInsert = data.hasOwnProperty(key) ? data[key] : match;
            if (typeof valueToInsert === 'string') {
                const trimmedValue = valueToInsert.trim();
                if (trimmedValue.startsWith('<p>') && trimmedValue.endsWith('</p>')) {
                    const innerHtml = trimmedValue.substring(3, trimmedValue.length - 4);
                    if (!innerHtml.includes('<p>') && !innerHtml.includes('<div>')) valueToInsert = innerHtml;
                }
            }
            return valueToInsert;
        });
    }
    return result;
};

const renderCCTemplate = (recipients, data) => {
    if (!recipients || recipients.length === 0) return '';
    return recipients.map(recipient => renderTemplate(recipient, data)).join(';');
}

function generatePayload() {
    const fromTemplate = document.getElementById('email-from').value;
    const subjectTemplate = document.getElementById('email-subject').value;
    const bodyTemplate = quill.root.innerHTML;

    return studentDataCache.map(student => ({
        from: renderTemplate(fromTemplate, student),
        to: student.StudentEmail || '',
        cc: renderCCTemplate(ccRecipients, student),
        subject: renderTemplate(subjectTemplate, student),
        body: renderTemplate(bodyTemplate, student)
    })).filter(email => email.to && email.from);
}

async function executeSend() {
    document.getElementById('send-confirm-modal').classList.add('hidden');
    const status = document.getElementById('status');
    status.textContent = `Sending ${studentDataCache.length} emails...`;
    status.style.color = 'gray';

    const payload = generatePayload();
    lastSentPayload = payload;

    if(payload.length === 0) {
        status.textContent = 'No students with valid "To" and "From" email addresses found.';
        status.style.color = 'orange';
        return;
    }

    try {
        const response = await fetch(powerAutomateConnection.url, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(payload)
        });
        if (!response.ok) throw new Error(`HTTP error! status: ${response.status}`);
        status.textContent = `Successfully sent ${payload.length} emails!`;
        status.style.color = 'green';
        modalManager.showSuccessModal(payload.length, lastSentPayload, quill.root.innerHTML);
    } catch (error) {
        status.textContent = `Failed to send emails: ${error.message}`;
        status.style.color = 'red';
        console.error("Error sending emails:", error);
    }
}

function isValidHttpUrl(string) {
    try {
        const url = new URL(string);
        return url.protocol === "http:" || url.protocol === "https:";
    } catch (_) {
        return false;
    }
}

async function loadCustomParameters() {
    customParameters = await getCustomParameters();
    return customParameters;
}

async function getCustomParameters() {
    return Excel.run(async (context) => {
        const settings = context.workbook.settings;
        const paramsSetting = settings.getItemOrNullObject(CUSTOM_PARAMS_KEY);
        paramsSetting.load("value");
        await context.sync();
        return paramsSetting.value ? JSON.parse(paramsSetting.value) : [];
    });
}

async function saveCustomParameters(params) {
    await Excel.run(async (context) => {
        context.workbook.settings.add(CUSTOM_PARAMS_KEY, JSON.stringify(params));
        await context.sync();
    });
}

async function getTemplates() {
    return Excel.run(async (context) => {
        const settings = context.workbook.settings;
        const templatesSetting = settings.getItemOrNullObject(EMAIL_TEMPLATES_KEY);
        templatesSetting.load("value");
        await context.sync();
        return templatesSetting.value ? JSON.parse(templatesSetting.value) : [];
    });
}

async function saveTemplates(templates) {
    await Excel.run(async (context) => {
        context.workbook.settings.add(EMAIL_TEMPLATES_KEY, JSON.stringify(templates));
        await context.sync();
    });
}

function setupFromInput() {
    const container = document.getElementById('email-from-container');
    const input = document.getElementById('email-from-input');
    container.onclick = () => { input.focus(); lastFocusedInput = input; };
    input.onfocus = () => { lastFocusedInput = input; };
    input.onkeydown = (e) => {
        if (e.key === ',' || e.key === 'Enter' || e.key === ';') {
            e.preventDefault();
            addFromRecipient(input.value.trim());
            input.value = '';
        } else if (e.key === 'Backspace' && input.value === '' && fromPill.length > 0) {
            removeFromRecipient(0);
        }
    };
    input.onblur = () => { addFromRecipient(input.value.trim()); input.value = ''; };
}

function addFromRecipient(text) {
    if (text) {
        fromPill = [text]; // Can only have one "From" address
        renderFromPills();
    }
}

function removeFromRecipient(index) {
    fromPill.splice(index, 1);
    renderFromPills();
}

function renderFromPills() {
    const container = document.getElementById('email-from-container');
    const input = document.getElementById('email-from-input');
    const hiddenInput = document.getElementById('email-from');

    container.querySelectorAll('.pill').forEach(pill => pill.remove());
    
    fromPill.forEach((recipient, index) => {
        const pill = document.createElement('span');
        pill.className = recipient.startsWith('{') && recipient.endsWith('}') ? 'pill param' : 'pill';
        pill.textContent = recipient;
        const removeBtn = document.createElement('span');
        removeBtn.textContent = '×';
        removeBtn.className = 'pill-remove';
        removeBtn.onclick = (e) => { e.stopPropagation(); removeFromRecipient(index); };
        pill.appendChild(removeBtn);
        container.insertBefore(pill, input);
    });

    if (hiddenInput) {
        hiddenInput.value = fromPill.length > 0 ? fromPill[0] : '';
    }
    validateAllFields();
}

function setupCcInput() {
    const container = document.getElementById('email-cc-container');
    const input = document.getElementById('email-cc-input');
    container.onclick = () => { input.focus(); lastFocusedInput = input; };
    input.onfocus = () => { lastFocusedInput = input; };
    input.onkeydown = (e) => {
        if (e.key === ',' || e.key === 'Enter' || e.key === ';') {
            e.preventDefault();
            addCcRecipient(input.value.trim());
            input.value = '';
        } else if (e.key === 'Backspace' && input.value === '' && ccRecipients.length > 0) {
            removeCcRecipient(ccRecipients.length - 1);
        }
    };
    input.onblur = () => { addCcRecipient(input.value.trim()); input.value = ''; };
}

function addCcRecipient(text) {
    if (text) {
        ccRecipients.push(text);
        renderCCPills();
    }
}

function removeCcRecipient(index) {
    ccRecipients.splice(index, 1);
    renderCCPills();
}

function renderCCPills() {
    const container = document.getElementById('email-cc-container');
    const input = document.getElementById('email-cc-input');
    container.querySelectorAll('.pill').forEach(pill => pill.remove());
    ccRecipients.forEach((recipient, index) => {
        const pill = document.createElement('span');
        pill.className = recipient.startsWith('{') && recipient.endsWith('}') ? 'pill param' : 'pill';
        pill.textContent = recipient;
        const removeBtn = document.createElement('span');
        removeBtn.textContent = '×';
        removeBtn.className = 'pill-remove';
        removeBtn.onclick = (e) => { e.stopPropagation(); removeCcRecipient(index); };
        pill.appendChild(removeBtn);
        container.insertBefore(pill, input);
    });
}

function validateAllFields() {
    const from = document.getElementById('email-from').value;
    const subject = document.getElementById('email-subject').value;
    const body = quill.getText().trim();
    const recipientsSet = recipientSelection.hasBeenSet;

    const isFromValid = from && from.trim() !== '';
    const isSubjectValid = subject && subject.trim() !== '';
    const isBodyValid = body !== '';
    const areRecipientsValid = recipientsSet && studentDataCache.length > 0;

    const sendButton = document.getElementById('send-email-button');
    const tooltip = document.getElementById('send-button-tooltip');
    
    if (isFromValid && isSubjectValid && isBodyValid && areRecipientsValid) {
        sendButton.disabled = false;
        tooltip.style.visibility = 'hidden';
    } else {
        sendButton.disabled = true;
        tooltip.style.visibility = 'visible';
        let missing = [];
        if (!isFromValid) missing.push('From address');
        if (!areRecipientsValid) missing.push('Recipients');
        if (!isSubjectValid) missing.push('Subject');
        if (!isBodyValid) missing.push('Body');
        tooltip.textContent = `Required: ${missing.join(', ')}.`;
    }
}

async function preCacheRecipientData() {
    try {
        const ldaSelection = { type: 'lda', customSheetName: '', excludeDNC: true, excludeFillColor: true };
        const ldaResult = await _getStudentDataCore(ldaSelection);
        recipientDataCache.set('lda', ldaResult);

        const masterSelection = { type: 'master', customSheetName: '', excludeDNC: true, excludeFillColor: true };
        const masterResult = await _getStudentDataCore(masterSelection);
        recipientDataCache.set('master', masterResult);
    } catch (error) {
        console.warn("Pre-caching failed. This may happen if sheets are not yet created. The add-in will function normally.", error);
    }
}
