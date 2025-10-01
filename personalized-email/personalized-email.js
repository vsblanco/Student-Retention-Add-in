// V-7.0 - 2025-10-01 - 3:16 PM EDT
import { findColumnIndex, getTodaysLdaSheetName, getNameParts } from './utils.js';
import { EMAIL_TEMPLATES_KEY, CUSTOM_PARAMS_KEY, standardParameters, QUILL_EDITOR_CONFIG, COLUMN_MAPPINGS, PARAMETER_BUTTON_STYLES } from './constants.js';
import ModalManager from './modal.js';

let powerAutomateConnection = null;
let studentDataCache = [];
let recipientCountCache = { lda: null, master: null };
let lastFocusedInput = null;
let quill; // To hold the editor instance
let fromPill = [];
let ccRecipients = [];
let customParameters = [];
let modalManager;
let worksheetDataCache = {}; // Cache for worksheet data
let recipientSelection = { type: 'lda', customSheetName: '', excludeDNC: true, excludeFillColor: true };

/**
 * Checks if all required fields are filled and enables/disables the send button accordingly.
 */
function updateSendButtonState() {
    const sendButton = document.getElementById('send-email-button');
    const fromFilled = fromPill.length > 0;
    const recipientsSelected = studentDataCache.length > 0;
    const subjectFilled = document.getElementById('email-subject').value.trim() !== '';
    // The editor is considered non-empty if it contains more than just the initial newline character.
    const bodyFilled = quill.getLength() > 1;

    if (fromFilled && recipientsSelected && subjectFilled && bodyFilled) {
        sendButton.disabled = false;
        sendButton.classList.remove('opacity-50', 'cursor-not-allowed');
        sendButton.title = ''; // Clear the title when enabled
    } else {
        sendButton.disabled = true;
        sendButton.classList.add('opacity-50', 'cursor-not-allowed');
        const missing = [];
        if (!fromFilled) missing.push('From');
        if (!recipientsSelected) missing.push('Recipients');
        if (!subjectFilled) missing.push('Subject');
        if (!bodyFilled) missing.push('Body');
        sendButton.title = `Please fill in the following fields: ${missing.join(', ')}`;
    }
}


/**
 * The core data fetching function. It reads data from the specified sheet,
 * processes it, and returns a processed array of student data.
 * @param {object} selection - The recipient selection object.
 * @returns {Promise<Array>} A promise that resolves with the student data array.
 */
async function _getStudentDataCore(selection) {
    console.log("--- Starting Data Fetch & DNC Exclusion Process ---");
    const { type, customSheetName, excludeDNC, excludeFillColor } = selection;
    console.log(`Selection criteria: type=${type}, customSheet='${customSheetName}', excludeDNC=${excludeDNC}, excludeFillColor=${excludeFillColor}`);
    let sheetName;
    let processedStudents = [];

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
    
    worksheetDataCache = {}; 

    try {
        await Excel.run(async (context) => {
            // Step 1: Build a list of DNC students by their ID if exclusion is enabled.
            const dncStudentIdentifiers = new Set();
            const dncEntriesForLogging = []; // Array for detailed logging
            if (excludeDNC) {
                console.log("Step 1: Building DNC exclusion list...");
                try {
                    const historySheet = context.workbook.worksheets.getItem("Student History");
                    const historyRange = historySheet.getUsedRange();
                    historyRange.load("values");
                    await context.sync();
                    console.log("Successfully loaded 'Student History' sheet.");
                    
                    const historyValues = historyRange.values;
                    if (historyValues.length > 1) {
                        const historyHeaders = historyValues[0].map(h => String(h ?? '').toLowerCase());
                        const identifierIndex = findColumnIndex(historyHeaders, COLUMN_MAPPINGS.StudentIdentifier);
                        const tagsIndex = findColumnIndex(historyHeaders, COLUMN_MAPPINGS.Tags);
                        console.log(`'Student History' Headers found: StudentIdentifier at index ${identifierIndex}, Tags at index ${tagsIndex}`);

                        if (identifierIndex !== -1 && tagsIndex !== -1) {
                            for (let i = 1; i < historyValues.length; i++) {
                                const row = historyValues[i];
                                const tagsString = String(row[tagsIndex] || '').toUpperCase();
                                const individualTags = tagsString.split(',').map(t => t.trim());

                                const hasExcludableDnc = individualTags.some(tag => {
                                    const isPhoneDnc = tag === 'DNC - PHONE' || tag === 'DNC - OTHER PHONE';
                                    return tag.includes('DNC') && !isPhoneDnc;
                                });

                                if (hasExcludableDnc) {
                                    const studentIdentifier = row[identifierIndex];
                                    if (studentIdentifier) {
                                        const idStr = String(studentIdentifier);
                                        dncStudentIdentifiers.add(idStr);
                                        dncEntriesForLogging.push({ id: idStr, tags: row[tagsIndex] });
                                    }
                                }
                            }
                        } else {
                            console.warn("Could not find 'Student Identifier' or 'Tags' column in 'Student History'. Cannot perform DNC exclusion.");
                        }
                    }
                    console.log(`Finished building exclusion list. Found ${dncStudentIdentifiers.size} unique students with DNC tags.`);
                    console.log("DNC Entries found in 'Student History':", dncEntriesForLogging);
                } catch (error) {
                    console.error("Error processing 'Student History' sheet for DNC exclusion. Proceeding without it.", error);
                }
            } else {
                 console.log("DNC exclusion is turned OFF. Skipping exclusion list build.");
            }

            // Step 2: Fetch and process the main student list.
            console.log(`Step 2: Fetching recipients from '${sheetName}' sheet.`);
            const sheet = context.workbook.worksheets.getItem(sheetName);
            const usedRange = sheet.getUsedRange();

            const propertiesToLoad = {
                format: {
                    fill: {
                        color: true
                    }
                }
            };
            const cellProperties = usedRange.getCellProperties(propertiesToLoad);
            usedRange.load("values");
            
            await context.sync();
            console.log(`Successfully loaded '${sheetName}' sheet with values and cell properties.`);

            const values = usedRange.values;
            const formats = cellProperties.value; 
            
            const headers = values[0].map(h => String(h ?? '').toLowerCase());
            
            const colIndices = {};
            for (const key in COLUMN_MAPPINGS) {
                colIndices[key] = findColumnIndex(headers, COLUMN_MAPPINGS[key]);
            }
            console.log("Recipient sheet column indices found:", colIndices);
            
            await loadCustomParameters();

            const customParamIndices = {};
            customParameters.forEach(param => {
                const headerIndex = headers.indexOf(param.sourceColumn.toLowerCase());
                if (headerIndex !== -1) {
                    customParamIndices[param.name] = headerIndex;
                }
            });

            console.log("Step 3: Processing and filtering recipient list...");
            if (excludeFillColor && colIndices.Outreach !== -1) {
                console.log("--- Begin Outreach Fill Color Analysis ---");
            }

            for (let i = 1; i < values.length; i++) {
                const row = values[i];
                if (!row) continue; 

                const studentIdentifier = row[colIndices.StudentIdentifier];

                if (excludeDNC && colIndices.StudentIdentifier !== -1) {
                    if (studentIdentifier && dncStudentIdentifiers.has(String(studentIdentifier))) {
                        console.log(`Excluding student (name: ${row[colIndices.StudentName]}, ID: ${studentIdentifier}) because they are on the DNC list.`);
                        continue;
                    }
                }

                if (excludeFillColor && colIndices.Outreach !== -1) {
                    const cellFormat = formats[i] ? formats[i][colIndices.Outreach] : null;
                    const cellColor = cellFormat ? cellFormat.format.fill.color : '#FFFFFF';
                    
                    console.log(`- Student: ${row[colIndices.StudentName] || 'Unknown Name'}, Outreach Color: ${cellColor}`);

                    if (cellColor && cellColor !== '#FFFFFF' && cellColor !== '#000000') {
                        console.log(`  ↳ EXCLUDING student (name: ${row[colIndices.StudentName]}, ID: ${studentIdentifier}) because their Outreach cell has a fill color.`);
                        continue;
                    }
                }
                
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
                processedStudents.push(student);
            }
            if (excludeFillColor && colIndices.Outreach !== -1) {
                console.log("--- End Outreach Fill Color Analysis ---");
            }
        });
        console.log(`--- Process Complete for ${sheetName}. Final count: ${processedStudents.length} ---`);
        return processedStudents;
    } catch (error) {
        console.error(`--- A critical error occurred during data fetch for ${sheetName} ---`, error);
        if (error.code === 'ItemNotFound') {
            error.userFacingMessage = `Error: Sheet "${sheetName}" not found.`;
        }
        throw error;
    }
}

/**
 * In the background, fetches the counts for the two primary sheets (LDA and Master)
 * to make the recipient modal feel faster.
 */
async function preCacheRecipientCounts() {
    console.log("--- Starting background pre-cache of recipient counts ---");
    // Use the default exclusion settings for the pre-cache.
    const ldaSelection = { type: 'lda', customSheetName: '', excludeDNC: true, excludeFillColor: true };
    const masterSelection = { type: 'master', customSheetName: '', excludeDNC: true, excludeFillColor: true };

    try {
        const ldaStudents = await _getStudentDataCore(ldaSelection);
        recipientCountCache.lda = ldaStudents.length;
        console.log(`Pre-cached LDA count: ${recipientCountCache.lda}`);
    } catch (error) {
        recipientCountCache.lda = -1; // Use -1 to indicate an error (e.g., sheet not found)
        console.warn("Could not pre-cache LDA count:", error.message);
    }

    try {
        const masterStudents = await _getStudentDataCore(masterSelection);
        recipientCountCache.master = masterStudents.length;
        console.log(`Pre-cached Master List count: ${recipientCountCache.master}`);
    } catch (error) {
        recipientCountCache.master = -1; // Use -1 to indicate an error
        console.warn("Could not pre-cache Master List count:", error.message);
    }
    console.log("--- Background pre-cache complete ---");
}


Office.onReady((info) => {
    if (info.host === Office.HostType.Excel) {
        quill = new Quill('#editor-container', QUILL_EDITOR_CONFIG);

        /**
         * Populates the main studentDataCache with the final selection of students
         * after the user confirms their choice in the recipient modal.
         * @param {object} selection The final recipient selection from the modal.
         */
        async function getStudentDataWithUI(selection) {
            const status = document.getElementById('status');
            status.textContent = 'Loading selected students...';
            status.style.color = 'gray';
            try {
                // The core function is called here to populate the main cache for the app
                studentDataCache = await _getStudentDataCore(selection);
                status.textContent = `Loaded ${studentDataCache.length} students.`;
                setTimeout(() => status.textContent = '', 3000);
            } catch (error) {
                studentDataCache = []; // Clear cache on error
                const message = error.userFacingMessage || (error.userFacing ? error.message : 'An error occurred while loading students.');
                status.textContent = message;
                status.style.color = 'red';
                throw error; // Rethrow to be caught by the modal if needed
            }
            updateSendButtonState();
        }
        
        const appContext = {
            quill,
            getStudentDataWithUI,
            getStudentDataCore: _getStudentDataCore,
            updateRecipientSelection: (newSelection, count) => {
                recipientSelection = newSelection;
                const button = document.getElementById('select-students-button');
                if (count >= 0) {
                    button.textContent = `${count} Student${count !== 1 ? 's' : ''} Selected`;
                    button.classList.add('bg-green-100', 'text-green-800', 'font-semibold');
                } else {
                    button.textContent = 'Select Students';
                    button.classList.remove('bg-green-100', 'text-green-800', 'font-semibold');
                }
            },
            recipientSelection,
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
            get studentDataCache() { return studentDataCache; },
            get recipientCountCache() { return recipientCountCache; },
            preCacheRecipientCounts
        };

        modalManager = new ModalManager(appContext);
        
        // Wire up button event handlers
        document.getElementById("send-email-button").onclick = () => modalManager.showSendConfirmModal();
        document.getElementById("create-connection-button").onclick = createConnection;
        document.getElementById("select-students-button").onclick = () => modalManager.showRecipientModal();
        
        setupFromInput();
        setupCcInput();
        setupExampleContextMenu();
        
        const subjectInput = document.getElementById('email-subject');
        subjectInput.addEventListener('focus', () => lastFocusedInput = subjectInput);
        subjectInput.addEventListener('input', updateSendButtonState);
        
        quill.on('selection-change', (range) => {
            if (range) lastFocusedInput = quill;
        });
        quill.on('text-change', updateSendButtonState);

        loadCustomParameters().then(populateParameterButtons);
        checkConnection().then(() => {
             if (powerAutomateConnection) {
                preCacheRecipientCounts(); // Initial cache on load
            }
        });

        updateSendButtonState(); // Call once on load to set initial disabled state
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
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(payload)
        });
        if (!response.ok) throw new Error(`HTTP error! status: ${response.status}`);
        status.textContent = `Successfully sent ${payload.length} emails!`;
        status.style.color = 'green';
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

    // Sync with the hidden input for modal.js
    if (hiddenInput) {
        hiddenInput.value = fromPill.length > 0 ? fromPill[0] : '';
    }
    updateSendButtonState();
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

