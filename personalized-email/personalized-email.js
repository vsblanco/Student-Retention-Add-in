// V-3.6 - 2025-09-17 - 4:00 PM EDT
import { findColumnIndex, getTodaysLdaSheetName, getNameParts } from './utils.js';
import { EMAIL_TEMPLATES_KEY, CUSTOM_PARAMS_KEY, standardParameters, QUILL_EDITOR_CONFIG, COLUMN_MAPPINGS, PARAMETER_BUTTON_STYLES } from './constants.js';
import ModalManager from './modal.js';

let powerAutomateConnection = null;
let studentDataCache = [];
let lastFocusedInput = null;
let quill; // To hold the editor instance
let ccRecipients = [];
let customParameters = [];
let modalManager;
let worksheetDataCache = {}; // Cache for worksheet data

Office.onReady((info) => {
    if (info.host === Office.HostType.Excel) {
        quill = new Quill('#editor-container', QUILL_EDITOR_CONFIG);
        
        const appContext = {
            quill,
            getStudentData: async () => { 
                await getStudentData(); // Ensures cache is fresh
                return studentDataCache;
            },
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
        };

        modalManager = new ModalManager(appContext);

        // Main Buttons not related to modals
        document.getElementById("send-email-button").onclick = () => modalManager.showSendConfirmModal();
        document.getElementById("create-connection-button").onclick = createConnection;

        // Dropdown listener
        document.getElementById('recipient-list').onchange = toggleCustomSheetInput;
        
        setupCcInput();
        const subjectInput = document.getElementById('email-subject');
        const fromInput = document.getElementById('email-from');
        
        // Track last focused element to insert parameters correctly
        subjectInput.addEventListener('focus', () => lastFocusedInput = subjectInput);
        fromInput.addEventListener('focus', () => lastFocusedInput = fromInput);
        quill.on('selection-change', (range) => {
            if (range) {
                lastFocusedInput = quill;
            }
        });

        loadCustomParameters().then(populateParameterButtons);
        checkConnection();
        toggleCustomSheetInput(); // Set initial state
    }
});

async function populateParameterButtons() {
    const standardContainer = document.getElementById('standard-parameter-buttons');
    const customContainer = document.getElementById('custom-parameter-buttons');
    const customSection = document.getElementById('custom-parameter-section');
    
    standardContainer.innerHTML = ''; 
    customContainer.innerHTML = '';

    const createButton = (param) => {
        const button = document.createElement('button');
        const isCustom = typeof param === 'object';
        const paramName = isCustom ? param.name : param;
        
        if (isCustom) {
            const hasMappings = param.mappings && param.mappings.length > 0;
            const hasNested = hasMappings && param.mappings.some(m => /\{(\w+)\}/.test(m.then));
            
            if (hasNested) {
                button.className = PARAMETER_BUTTON_STYLES.nested;
            } else if (hasMappings) {
                button.className = PARAMETER_BUTTON_STYLES.mapped;
            } else {
                button.className = PARAMETER_BUTTON_STYLES.custom;
            }
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
        customParameters.forEach(param => {
            customContainer.appendChild(createButton(param));
        });
        customSection.classList.remove('hidden');
    } else {
        customSection.classList.add('hidden');
    }
}


function insertParameter(param) {
    if (lastFocusedInput instanceof Quill) {
        const range = lastFocusedInput.getSelection(true);
        lastFocusedInput.insertText(range.index, param, 'user');
    } else if (lastFocusedInput && lastFocusedInput.id === 'email-cc-input') {
        addCcRecipient(param);
    } else if (lastFocusedInput) { // It's the subject or from input
        const start = lastFocusedInput.selectionStart;
        const end = lastFocusedInput.selectionEnd;
        const text = lastFocusedInput.value;
        lastFocusedInput.value = text.substring(0, start) + param + text.substring(end);
        lastFocusedInput.focus();
        lastFocusedInput.selectionStart = lastFocusedInput.selectionEnd = start + param.length;
    } else {
        // Default to editor if nothing is focused
        quill.focus();
        const length = quill.getLength();
        quill.insertText(length, param, 'user');
    }
}

function toggleCustomSheetInput() {
    const recipientList = document.getElementById('recipient-list');
    const customSheetContainer = document.getElementById('custom-sheet-container');
    if (recipientList.value === 'custom') {
        customSheetContainer.classList.remove('hidden');
    } else {
        customSheetContainer.classList.add('hidden');
    }
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

    const cellNum = parseFloat(cellValue);
    const conditionNum = parseFloat(mapping.if);

    switch (mapping.operator) {
        case 'eq': return cellStr === conditionStr;
        case 'neq': return cellStr !== conditionStr;
        case 'contains': return cellStr.includes(conditionStr);
        case 'does_not_contain': return !cellStr.includes(conditionStr);
        case 'starts_with': return cellStr.startsWith(conditionStr);
        case 'ends_with': return cellStr.endsWith(conditionStr);
        case 'gt': return !isNaN(cellNum) && !isNaN(conditionNum) && cellNum > conditionNum;
        case 'lt': return !isNaN(cellNum) && !isNaN(conditionNum) && cellNum < conditionNum;
        case 'gte': return !isNaN(cellNum) && !isNaN(conditionNum) && cellNum >= conditionNum;
        case 'lte': return !isNaN(cellNum) && !isNaN(conditionNum) && cellNum <= conditionNum;
        default: return false;
    }
}

/**
 * Fetches data for a given worksheet and caches it.
 * @param {string} sheetNameToFetch The name of the worksheet to get data from.
 * @returns {Promise<object|null>} The cached data object or null if not found.
 */
async function getWorksheetData(sheetNameToFetch) {
    if (worksheetDataCache[sheetNameToFetch]) {
        return worksheetDataCache[sheetNameToFetch];
    }

    try {
        await Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getItem(sheetNameToFetch);
            const range = sheet.getUsedRange();
            range.load("values");
            await context.sync();

            const values = range.values;
            if (values.length > 1) {
                worksheetDataCache[sheetNameToFetch] = {
                    headers: values[0].map(h => String(h ?? '').toLowerCase()),
                    values: values.slice(1)
                };
            } else {
                 worksheetDataCache[sheetNameToFetch] = { headers: [], values: [] };
            }
        });
        return worksheetDataCache[sheetNameToFetch];
    } catch (error) {
        if (error.code !== 'ItemNotFound') {
            console.error(`Error loading '${sheetNameToFetch}' sheet:`, error);
        }
        worksheetDataCache[sheetNameToFetch] = null; // Cache null if not found
        return null;
    }
}


async function getStudentData() {
    const recipientListValue = document.getElementById('recipient-list').value;
    const status = document.getElementById('status');
    let sheetName;

    if (recipientListValue === 'custom') {
        sheetName = document.getElementById('custom-sheet-name').value.trim();
        if (!sheetName) {
            status.textContent = 'Please enter a custom sheet name.';
            status.style.color = 'red';
            throw new Error('Custom sheet name is required.');
        }
    } else {
        sheetName = recipientListValue === 'lda' ? getTodaysLdaSheetName() : 'Master List';
    }
    
    status.textContent = `Fetching students from "${sheetName}"...`;
    status.style.color = 'gray';
    
    studentDataCache = [];
    worksheetDataCache = {}; // Reset cache for each run

    try {
        await Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getItem(sheetName);
            const usedRange = sheet.getUsedRange();
            usedRange.load("values");
            await context.sync();

            const values = usedRange.values;
            const headers = values[0].map(h => String(h ?? '').toLowerCase());
            
            await loadCustomParameters();

            const colIndices = {};
            for (const key in COLUMN_MAPPINGS) {
                colIndices[key] = findColumnIndex(headers, COLUMN_MAPPINGS[key]);
            }

            const customParamIndices = {};
            customParameters.forEach(param => {
                const headerIndex = headers.indexOf(param.sourceColumn.toLowerCase());
                if (headerIndex !== -1) {
                    customParamIndices[param.name] = headerIndex;
                }
            });

            for (let i = 1; i < values.length; i++) {
                const row = values[i];
                const studentName = row[colIndices.StudentName] ?? '';
                const nameParts = getNameParts(studentName);

                const student = {
                    StudentName: studentName,
                    FirstName: nameParts.first,
                    LastName: nameParts.last,
                    StudentEmail: row[colIndices.StudentEmail] ?? '',
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
                            const argValues = [getWorksheetData, '']; // Placeholder for sourceColumnValue

                            let userScript = param.script;
                            const mainSourceColIndex = headers.indexOf(param.sourceColumn.toLowerCase());
                            if (mainSourceColIndex !== -1) {
                                argValues[1] = row[mainSourceColIndex];
                            }

                            if (param.scriptInputs) {
                                for (const varName in param.scriptInputs) {
                                    const sourceColName = param.scriptInputs[varName];
                                    const sourceColIndex = headers.indexOf(sourceColName.toLowerCase());
                                    const varValue = (sourceColIndex !== -1) ? row[sourceColIndex] : undefined;
                                    
                                    argNames.push(varName);
                                    argValues.push(varValue);

                                    const declarationRegex = new RegExp(`\\blet\\s+${varName}\\s*;`, 'g');
                                    userScript = userScript.replace(declarationRegex, '');
                                }
                            }
                
                            let finalScriptBody;
                            const isAsync = /\bawait\b/.test(userScript);
                            const hasReturn = /\breturn\b/.test(userScript);

                            if (isAsync) {
                                finalScriptBody = hasReturn ? userScript : `return (async () => { ${userScript} })();`;
                                if (!hasReturn) console.warn(`Warning: Async script for parameter "${param.name}" is missing an explicit 'return' statement.`);
                            } else {
                                finalScriptBody = hasReturn ? userScript : `return (() => { "use strict"; ${userScript} })();`;
                            }

                            const executor = isAsync 
                                ? new Function(...argNames, `return (async () => { "use strict"; ${finalScriptBody} })();`)
                                : new Function(...argNames, finalScriptBody);
                            
                            value = await executor(...argValues);

                        } catch (e) {
                            console.error(`Error executing script for parameter "${param.name}":`, e);
                            value = `[SCRIPT ERROR]`;
                        }
                    } else { // Handle value-mapping logic
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
                            if (!mappingFound) {
                                value = cellValue;
                            }
                        }
                    }
                    student[param.name] = value;
                }
                
                studentDataCache.push(student);
            }
        });

        status.textContent = `Found ${studentDataCache.length} students.`;
        setTimeout(() => status.textContent = '', 3000);

    } catch (error) {
        if (error.code === 'ItemNotFound') {
            status.textContent = `Error: Sheet "${sheetName}" not found.`;
        } else {
            status.textContent = 'An error occurred while fetching data.';
        }
        status.style.color = 'red';
        console.error(error);
        throw error;
    }
}

const renderTemplate = (template, data) => {
    if (!template) return '';
    let result = template;
    let iterations = 0;
    const maxIterations = 10;
    const regex = /\{(\w+)\}/g;
    
    while (result.match(regex) && iterations < maxIterations) {
        result = result.replace(regex, (match, key) => {
            let valueToInsert = data.hasOwnProperty(key) ? data[key] : match;
            if (typeof valueToInsert === 'string') {
                const trimmedValue = valueToInsert.trim();
                if (trimmedValue.startsWith('<p>') && trimmedValue.endsWith('</p>')) {
                    const innerHtml = trimmedValue.substring(3, trimmedValue.length - 4);
                    if (!innerHtml.includes('<p>') && !innerHtml.includes('<div>')) {
                         valueToInsert = innerHtml;
                    }
                }
            }
            return valueToInsert;
        });
        iterations++;
    }

    return result;
};


const renderCCTemplate = (recipients, data) => {
    if (!recipients || recipients.length === 0) return '';
    return recipients.map(recipient => renderTemplate(recipient, data)).join(';');
}

async function executeSend() {
    document.getElementById('send-confirm-modal').classList.add('hidden');
    const status = document.getElementById('status');
    status.textContent = `Sending ${studentDataCache.length} emails...`;
    status.style.color = 'gray';

    const fromTemplate = document.getElementById('email-from').value;
    const subjectTemplate = document.getElementById('email-subject').value;
    const bodyTemplate = quill.root.innerHTML;

    const payload = studentDataCache.map(student => ({
        from: renderTemplate(fromTemplate, student),
        to: student.StudentEmail || '',
        cc: renderCCTemplate(ccRecipients, student),
        subject: renderTemplate(subjectTemplate, student),
        body: renderTemplate(bodyTemplate, student)
    })).filter(email => email.to && email.from);

    if(payload.length === 0) {
        status.textContent = 'No students with valid "To" and "From" email addresses found.';
        status.style.color = 'orange';
        return;
    }

    try {
        const response = await fetch(powerAutomateConnection.url, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify(payload)
        });

        if (!response.ok) {
            throw new Error(`HTTP error! status: ${response.status}`);
        }

        status.textContent = `Successfully sent ${payload.length} emails!`;
        status.style.color = 'green';
    } catch (error) {
        status.textContent = `Failed to send emails: ${error.message}`;
        status.style.color = 'red';
        console.error("Error sending emails:", error);
    }
}


function isValidHttpUrl(string) {
    let url;
    try {
        url = new URL(string);
    } catch (_) {
        return false;
    }
    return url.protocol === "http:" || url.protocol === "https:";
}

// --- Custom Parameter Functions ---

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
        const settings = context.workbook.settings;
        settings.add(CUSTOM_PARAMS_KEY, JSON.stringify(params));
        await context.sync();
    });
}


// --- Template Functions --- 
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
        const settings = context.workbook.settings;
        settings.add(EMAIL_TEMPLATES_KEY, JSON.stringify(templates));
        await context.sync();
    });
}


// --- CC Pillbox Functions ---
function setupCcInput() {
    const container = document.getElementById('email-cc-container');
    const input = document.getElementById('email-cc-input');

    container.addEventListener('click', () => {
        input.focus();
        lastFocusedInput = input;
    });

    input.addEventListener('focus', () => lastFocusedInput = input);

    input.addEventListener('keydown', (e) => {
        if (e.key === ',' || e.key === 'Enter' || e.key === ';') {
            e.preventDefault();
            addCcRecipient(input.value.trim());
            input.value = '';
        } else if (e.key === 'Backspace' && input.value === '') {
            if (ccRecipients.length > 0) {
                removeCcRecipient(ccRecipients.length - 1);
            }
        }
    });

    input.addEventListener('blur', () => {
        addCcRecipient(input.value.trim());
        input.value = '';
    });
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
    
    container.querySelectorAll('.cc-pill').forEach(pill => pill.remove());

    ccRecipients.forEach((recipient, index) => {
        const pill = document.createElement('span');
        const isParam = recipient.startsWith('{') && recipient.endsWith('}');
        pill.className = isParam ? 'cc-pill param' : 'cc-pill';
        pill.textContent = recipient;
        
        const removeBtn = document.createElement('span');
        removeBtn.textContent = '×';
        removeBtn.className = 'cc-pill-remove';
        removeBtn.onclick = (e) => {
            e.stopPropagation();
            removeCcRecipient(index);
        };
        
        pill.appendChild(removeBtn);
        container.insertBefore(pill, input);
    });
}

