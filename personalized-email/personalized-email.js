import { findColumnIndex, getTodaysLdaSheetName, getNameParts } from './utils.js';

let powerAutomateConnection = null;
let studentDataCache = [];
let lastFocusedInput = null;
let quill; // To hold the editor instance
let ccRecipients = [];
let specialParameters = [];

const standardParameters = ['FirstName', 'LastName', 'StudentName', 'StudentEmail', 'PersonalEmail', 'Grade', 'DaysOut'];
const EMAIL_TEMPLATES_KEY = "emailTemplates";
const CUSTOM_PARAMS_KEY = "customEmailParameters";

const PAYLOAD_SCHEMA = {
    "type": "array",
    "items": {
        "type": "object",
        "properties": {
            "from": { "type": "string" },
            "to": { "type": "string" },
            "cc": { "type": "string" },
            "subject": { "type": "string" },
            "body": { "type": "string" }
        },
        "required": ["from", "to", "subject", "body"]
    }
};


Office.onReady((info) => {
    if (info.host === Office.HostType.Excel) {
        // Main Buttons
        document.getElementById("send-email-button").onclick = sendEmail;
        document.getElementById("create-connection-button").onclick = createConnection;
        document.getElementById('show-example-button').onclick = showExample;
        document.getElementById('show-payload-button').onclick = showPayload;
        document.getElementById('templates-button').onclick = showTemplatesModal;
        document.getElementById('create-special-param-button').onclick = () => showSpecialParamModal();

        // Dropdown listener
        document.getElementById('recipient-list').onchange = toggleCustomSheetInput;

        // Modal Close Buttons
        document.getElementById('close-example-modal-button').onclick = () => document.getElementById('example-modal').classList.add('hidden');
        document.getElementById('close-payload-modal-button').onclick = () => document.getElementById('payload-modal').classList.add('hidden');
        document.getElementById('close-templates-modal-button').onclick = () => document.getElementById('templates-modal').classList.add('hidden');
        document.getElementById('cancel-save-template-button').onclick = () => document.getElementById('save-template-modal').classList.add('hidden');
        document.getElementById('cancel-send-button').onclick = () => document.getElementById('send-confirm-modal').classList.add('hidden');
        document.getElementById('cancel-special-param-button').onclick = () => document.getElementById('special-param-modal').classList.add('hidden');
        document.getElementById('close-manage-params-button').onclick = () => document.getElementById('manage-params-modal').classList.add('hidden');
        
        // Modal Action Buttons
        document.getElementById('toggle-payload-schema-button').onclick = togglePayloadView;
        document.getElementById('save-current-template-button').onclick = showSaveTemplateModal;
        document.getElementById('confirm-save-template-button').onclick = saveTemplate;
        document.getElementById('confirm-send-button').onclick = executeSend;
        document.getElementById('save-special-param-button').onclick = saveSpecialParameter;
        document.getElementById('add-mapping-button').onclick = () => addMappingRow();
        document.getElementById('manage-special-params-button').onclick = showManageParamsModal;

        // Initialize Quill Editor
        quill = new Quill('#editor-container', {
            theme: 'snow',
            modules: {
                toolbar: [
                    ['bold', 'italic', 'underline'],
                    [{'list': 'ordered'}, {'list': 'bullet'}],
                    [{'color': []}, {'background': []}],
                    ['link']
                ]
            }
        });
        
        // --- START: Custom Blot for Parameters ---
        let Inline = Quill.import('blots/inline');

        class ParameterBlot extends Inline {
            static create(value) {
                let node = super.create();
                node.setAttribute('contenteditable', 'false');
                node.setAttribute('data-param', value.name);
                node.classList.add('parameter-tag');
                if (value.isSpecial) {
                    node.classList.add('special');
                }
                node.innerText = `{${value.name}}`;
                return node;
            }

            static formats(node) {
                return {
                    name: node.getAttribute('data-param'),
                    isSpecial: node.classList.contains('special')
                };
            }
        }
        ParameterBlot.blotName = 'parameter';
        ParameterBlot.tagName = 'SPAN';
        Quill.register(ParameterBlot);
        // --- END: Custom Blot for Parameters ---
        
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
    const specialContainer = document.getElementById('special-parameter-buttons');
    const specialSection = document.getElementById('special-parameters-section');

    standardContainer.innerHTML = '';
    specialContainer.innerHTML = '';

    standardParameters.forEach(param => {
        const button = document.createElement('button');
        button.className = 'px-2 py-1 bg-gray-200 text-gray-700 text-xs rounded hover:bg-gray-300';
        button.textContent = `{${param}}`;
        button.onclick = () => insertParameter(`{${param}}`);
        standardContainer.appendChild(button);
    });

    if (specialParameters.length > 0) {
        specialSection.classList.remove('hidden');
        specialParameters.forEach(param => {
            const button = document.createElement('button');
            button.className = 'px-2 py-1 bg-blue-100 text-blue-800 text-xs rounded hover:bg-blue-200';
            button.textContent = `{${param.name}}`;
            button.onclick = () => insertParameter(`{${param.name}}`);
            specialContainer.appendChild(button);
        });
    } else {
        specialSection.classList.add('hidden');
    }
}


function insertParameter(param) {
    const paramName = param.replace(/[{}]/g, ''); // Get name without braces
    const isSpecial = specialParameters.some(p => p.name === paramName);

    if (lastFocusedInput instanceof Quill) {
        const range = lastFocusedInput.getSelection(true);
        // Insert the custom blot
        lastFocusedInput.insertEmbed(range.index, 'parameter', { name: paramName, isSpecial: isSpecial }, Quill.sources.USER);
        // Move cursor after the inserted blot
        lastFocusedInput.setSelection(range.index + 1, Quill.sources.USER);
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
        quill.insertEmbed(length, 'parameter', { name: paramName, isSpecial: isSpecial }, Quill.sources.USER);
        quill.setSelection(length + 1, Quill.sources.USER);
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
    
    studentDataCache = []; // Clear cache before fetching

    await Excel.run(async (context) => {
        try {
            const sheet = context.workbook.worksheets.getItem(sheetName);
            const usedRange = sheet.getUsedRange();
            usedRange.load("values");
            await context.sync();

            const values = usedRange.values;
            const headers = values[0].map(h => String(h ?? '').toLowerCase());
            
            await loadCustomParameters();

            const colIndices = {
                StudentName: findColumnIndex(headers, ["studentname", "student name"]),
                StudentEmail: findColumnIndex(headers, ["student email", "school email", "email"]),
                PersonalEmail: findColumnIndex(headers, ["personal email", "otheremail"]),
                Grade: findColumnIndex(headers, ["grade", "course grade"]),
                DaysOut: findColumnIndex(headers, ["days out", "daysout"]),
            };

            const specialParamIndices = {};
            specialParameters.forEach(param => {
                const headerIndex = headers.indexOf(param.sourceColumn.toLowerCase());
                if (headerIndex !== -1) {
                    specialParamIndices[param.name] = headerIndex;
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
                };

                specialParameters.forEach(param => {
                    const colIndex = specialParamIndices[param.name];
                    let value = param.defaultValue ?? '';
                    if (colIndex !== undefined) {
                        const cellValue = row[colIndex];
                        let mappingFound = false;
                        if (param.mappings && cellValue != null) {
                            for (const mapping of param.mappings) {
                                if (String(cellValue).trim().toLowerCase() === String(mapping.if).trim().toLowerCase()) {
                                    value = mapping.then;
                                    mappingFound = true;
                                    break;
                                }
                            }
                        }
                        if (!mappingFound && cellValue != null) {
                             value = cellValue ?? param.defaultValue ?? '';
                        }
                    }
                    student[param.name] = value;
                });
                studentDataCache.push(student);
            }
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
    });
}

const renderTemplate = (template, data) => {
    if (!template) return '';
    return template.replace(/\{(\w+)\}/g, (match, key) => {
        return (data[key] ?? match);
    });
};

const renderBodyTemplate = (bodyHtml, data) => {
    if (!bodyHtml) return '';
    // Create a temporary DOM element to parse the HTML
    const tempDiv = document.createElement('div');
    tempDiv.innerHTML = bodyHtml;

    // Find all parameter tags
    tempDiv.querySelectorAll('.parameter-tag').forEach(tag => {
        const paramName = tag.getAttribute('data-param');
        const value = data[paramName] ?? tag.innerText; // Fallback to the tag text if data not found
        // Replace the tag with a simple text node
        tag.replaceWith(document.createTextNode(value));
    });

    return tempDiv.innerHTML;
};

const renderCCTemplate = (recipients, data) => {
    if (!recipients || recipients.length === 0) return '';
    return recipients.map(recipient => renderTemplate(recipient, data)).join(';');
}

async function showExample() {
    const status = document.getElementById('status');
    try {
        await getStudentData();

        if (studentDataCache.length === 0) {
            status.textContent = 'No students found to generate an example.';
            status.style.color = 'orange';
            return;
        }

        const randomStudent = studentDataCache[Math.floor(Math.random() * studentDataCache.length)];
        
        const fromTemplate = document.getElementById('email-from').value;
        const subjectTemplate = document.getElementById('email-subject').value;
        const bodyTemplate = quill.root.innerHTML;

        document.getElementById('example-from').textContent = renderTemplate(fromTemplate, randomStudent) || '[Not Specified]';
        document.getElementById('example-to').textContent = randomStudent.StudentEmail || '[No Email Found]';
        document.getElementById('example-cc').textContent = renderCCTemplate(ccRecipients, randomStudent) || '[Not Specified]';
        document.getElementById('example-subject').textContent = renderTemplate(subjectTemplate, randomStudent);
        document.getElementById('example-body').innerHTML = renderBodyTemplate(bodyTemplate, randomStudent);

        document.getElementById('example-modal').classList.remove('hidden');

    } catch (error) {
        // Error message is already set by getStudentData
    }
}

async function showPayload() {
    const status = document.getElementById('status');
    try {
        await getStudentData();

        if (studentDataCache.length === 0) {
            status.textContent = 'No students found to generate a payload.';
            status.style.color = 'orange';
            return;
        }

        const fromTemplate = document.getElementById('email-from').value;
        const subjectTemplate = document.getElementById('email-subject').value;
        const bodyTemplate = quill.root.innerHTML;

        const payload = studentDataCache.map(student => ({
            from: renderTemplate(fromTemplate, student),
            to: student.StudentEmail || '',
            cc: renderCCTemplate(ccRecipients, student),
            subject: renderTemplate(subjectTemplate, student),
            body: renderBodyTemplate(bodyTemplate, student)
        }));

        document.getElementById('payload-content').textContent = JSON.stringify(payload, null, 2);
        document.getElementById('schema-content').textContent = JSON.stringify(PAYLOAD_SCHEMA, null, 2);
        
        document.getElementById('payload-content').classList.remove('hidden');
        document.getElementById('schema-content').classList.add('hidden');
        document.getElementById('payload-modal-title').textContent = 'Request Payload';
        document.getElementById('toggle-payload-schema-button').textContent = 'Show Schema';

        document.getElementById('payload-modal').classList.remove('hidden');

    } catch (error) {
        // Error message is already set by getStudentData
    }
}

function togglePayloadView() {
    const payloadContent = document.getElementById('payload-content');
    const schemaContent = document.getElementById('schema-content');
    const title = document.getElementById('payload-modal-title');
    const button = document.getElementById('toggle-payload-schema-button');

    if (!payloadContent.classList.contains('hidden')) {
        payloadContent.classList.add('hidden');
        schemaContent.classList.remove('hidden');
        title.textContent = 'Request Body JSON Schema';
        button.textContent = 'Show Payload';
    } else {
        payloadContent.classList.remove('hidden');
        schemaContent.classList.add('hidden');
        title.textContent = 'Request Payload';
        button.textContent = 'Show Schema';
    }
}

async function sendEmail() {
    const status = document.getElementById('status');
    try {
        await getStudentData();

        if (studentDataCache.length === 0) {
            status.textContent = 'No students to send emails to.';
            status.style.color = 'orange';
            return;
        }
        
        const confirmMessage = document.getElementById('send-confirm-message');
        confirmMessage.textContent = `You are about to send emails to ${studentDataCache.length} student(s). Do you want to proceed?`;
        document.getElementById('send-confirm-modal').classList.remove('hidden');

    } catch (error) {
        // Error message is already set by getStudentData
    }
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
        body: renderBodyTemplate(bodyTemplate, student)
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

async function showTemplatesModal() {
    const container = document.getElementById('templates-list-container');
    container.innerHTML = '<p class="text-gray-500">Loading templates...</p>';
    document.getElementById('templates-modal').classList.remove('hidden');

    const templates = await getTemplates();
    container.innerHTML = '';
    if (templates.length === 0) {
        container.innerHTML = '<p class="text-gray-500 text-center">No saved templates found.</p>';
        return;
    }

    templates.forEach(template => {
        const div = document.createElement('div');
        div.className = 'p-3 border rounded-md bg-gray-50';
        div.innerHTML = `
            <div class="flex justify-between items-start">
                <div>
                    <p class="font-semibold text-gray-800">${template.name}</p>
                    <p class="text-xs text-gray-500">by ${template.author} on ${new Date(template.timestamp).toLocaleDateString()}</p>
                </div>
                <div class="flex gap-2">
                    <button data-id="${template.id}" class="load-template-btn px-3 py-1 bg-blue-100 text-blue-800 text-xs font-semibold rounded-md hover:bg-blue-200">Load</button>
                    <button data-id="${template.id}" class="delete-template-btn px-3 py-1 bg-red-100 text-red-800 text-xs font-semibold rounded-md hover:bg-red-200">Delete</button>
                </div>
            </div>
        `;
        container.appendChild(div);
    });
    
    container.querySelectorAll('.load-template-btn').forEach(btn => {
        btn.onclick = () => loadTemplate(btn.dataset.id);
    });
    container.querySelectorAll('.delete-template-btn').forEach(btn => {
        btn.onclick = () => deleteTemplate(btn.dataset.id);
    });
}

function showSaveTemplateModal() {
    document.getElementById('templates-modal').classList.add('hidden');
    document.getElementById('template-name').value = '';
    document.getElementById('template-author').value = '';
    document.getElementById('save-template-status').textContent = '';
    document.getElementById('save-template-modal').classList.remove('hidden');
}

async function saveTemplate() {
    const name = document.getElementById('template-name').value.trim();
    const author = document.getElementById('template-author').value.trim();
    const status = document.getElementById('save-template-status');

    if (!name || !author) {
        status.textContent = 'Name and Author are required.';
        status.style.color = 'red';
        return;
    }

    status.textContent = 'Saving...';
    status.style.color = 'gray';

    const newTemplate = {
        id: 'template_' + new Date().getTime(),
        name: name,
        author: author,
        timestamp: new Date().toISOString(),
        from: document.getElementById('email-from').value,
        subject: document.getElementById('email-subject').value,
        cc: ccRecipients,
        body: quill.root.innerHTML
    };

    const templates = await getTemplates();
    templates.push(newTemplate);
    await saveTemplates(templates);

    status.textContent = 'Template saved!';
    status.style.color = 'green';
    setTimeout(() => {
        document.getElementById('save-template-modal').classList.add('hidden');
    }, 1500);
}

async function loadTemplate(templateId) {
    const templates = await getTemplates();
    const template = templates.find(t => t.id === templateId);
    if (template) {
        document.getElementById('email-from').value = template.from || '';
        document.getElementById('email-subject').value = template.subject;
        ccRecipients = template.cc || [];
        renderCCPills();
        quill.root.innerHTML = template.body;
        document.getElementById('templates-modal').classList.add('hidden');
    }
}

async function deleteTemplate(templateId) {
    let templates = await getTemplates();
    templates = templates.filter(t => t.id !== templateId);
    await saveTemplates(templates);
    await showTemplatesModal();
}

// --- Special Parameter Functions ---

async function loadCustomParameters() {
    specialParameters = await getCustomParameters();
    return specialParameters;
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

function showSpecialParamModal(paramToEdit = null) {
    const modalTitle = document.getElementById('special-param-modal-title');
    const nameInput = document.getElementById('param-name');
    const sourceColumnInput = document.getElementById('param-source-column');
    const defaultValueInput = document.getElementById('param-default-value');
    const mappingContainer = document.getElementById('param-mapping-container');
    const editIdInput = document.getElementById('param-edit-id');

    mappingContainer.innerHTML = '';
    document.getElementById('save-param-status').textContent = '';

    if (paramToEdit) {
        modalTitle.textContent = 'Edit Special Parameter';
        editIdInput.value = paramToEdit.id;
        nameInput.value = paramToEdit.name;
        sourceColumnInput.value = paramToEdit.sourceColumn;
        defaultValueInput.value = paramToEdit.defaultValue;
        
        paramToEdit.mappings.forEach(m => addMappingRow(m.if, m.then));

    } else {
        modalTitle.textContent = 'Create Special Parameter';
        editIdInput.value = '';
        nameInput.value = '';
        sourceColumnInput.value = '';
        defaultValueInput.value = '';
    }
    
    document.getElementById('special-param-modal').classList.remove('hidden');
}

function addMappingRow(ifValue = '', thenValue = '') {
    const container = document.getElementById('param-mapping-container');
    const div = document.createElement('div');
    div.className = 'flex items-center gap-2 mapping-row';
    div.innerHTML = `
        <span class="text-sm text-gray-500">If cell is</span>
        <input type="text" class="mapping-if flex-1 px-2 py-1 border border-gray-300 rounded-md text-sm" placeholder="e.g., Bob" value="${ifValue}">
        <span class="text-sm text-gray-500">then value is</span>
        <input type="text" class="mapping-then flex-1 px-2 py-1 border border-gray-300 rounded-md text-sm" placeholder="e.g., bobjones@gmail.com" value="${thenValue}">
        <button class="remove-mapping-btn text-red-500 hover:text-red-700 text-lg">&times;</button>
    `;
    div.querySelector('.remove-mapping-btn').onclick = () => div.remove();
    container.appendChild(div);
}

async function saveSpecialParameter() {
    const status = document.getElementById('save-param-status');
    const nameInput = document.getElementById('param-name');
    const name = nameInput.value.trim();
    const paramId = document.getElementById('param-edit-id').value;

    // Validation
    if (!/^[a-zA-Z0-9]+$/.test(name)) {
        status.textContent = 'Name must be alphanumeric with no spaces.';
        status.style.color = 'red';
        return;
    }
    const existingParam = specialParameters.find(p => p.name.toLowerCase() === name.toLowerCase());
    if (standardParameters.includes(name) || (existingParam && existingParam.id !== paramId)) {
        status.textContent = 'This parameter name is already in use.';
        status.style.color = 'red';
        return;
    }

    const sourceColumn = document.getElementById('param-source-column').value;
    const defaultValue = document.getElementById('param-default-value').value.trim();
    
    const mappings = [];
    document.querySelectorAll('#param-mapping-container .mapping-row').forEach(row => {
        const ifValue = row.querySelector('.mapping-if').value.trim();
        const thenValue = row.querySelector('.mapping-then').value.trim();
        if (ifValue) {
            mappings.push({ if: ifValue, then: thenValue });
        }
    });

    const paramData = {
        name,
        sourceColumn,
        defaultValue,
        mappings
    };

    status.textContent = 'Saving...';
    status.style.color = 'gray';

    let currentParams = await getCustomParameters();
    if (paramId) { // Editing existing
        const index = currentParams.findIndex(p => p.id === paramId);
        if (index > -1) {
            currentParams[index] = { ...currentParams[index], ...paramData };
        }
    } else { // Creating new
        paramData.id = 'sparam_' + new Date().getTime();
        currentParams.push(paramData);
    }
    
    await saveCustomParameters(currentParams);
    
    await loadCustomParameters();
    await populateParameterButtons();

    status.textContent = 'Parameter saved successfully!';
    status.style.color = 'green';
    setTimeout(() => {
        document.getElementById('special-param-modal').classList.add('hidden');
    }, 1500);
}

async function showManageParamsModal() {
    document.getElementById('special-param-modal').classList.add('hidden');
    const listContainer = document.getElementById('manage-params-list');
    listContainer.innerHTML = '<p class="text-gray-500">Loading...</p>';
    document.getElementById('manage-params-modal').classList.remove('hidden');

    const params = await getCustomParameters();
    listContainer.innerHTML = '';
    if (params.length === 0) {
        listContainer.innerHTML = '<p class="text-gray-500 text-center">No special parameters created yet.</p>';
        return;
    }

    params.forEach(param => {
        const div = document.createElement('div');
        div.className = 'p-3 border-b';
        let mappingsHtml = param.mappings.map(m => `<div class="text-xs ml-4"><span class="text-gray-500">If '${m.if}' &rarr;</span> '${m.then}'</div>`).join('');
        if (!mappingsHtml) mappingsHtml = '<div class="text-xs ml-4 text-gray-400">No mappings</div>';

        div.innerHTML = `
            <div class="flex justify-between items-start">
                <div>
                    <p class="font-semibold text-gray-800">{${param.name}}</p>
                    <p class="text-xs text-gray-500">Reads from column: <strong>${param.sourceColumn}</strong></p>
                    <p class="text-xs text-gray-500">Default: <strong>${param.defaultValue || '<em>(none)</em>'}</strong></p>
                </div>
                <div class="flex gap-2">
                    <button data-id="${param.id}" class="duplicate-param-btn px-3 py-1 bg-gray-100 text-gray-800 text-xs font-semibold rounded-md hover:bg-gray-200">Duplicate</button>
                    <button data-id="${param.id}" class="edit-param-btn px-3 py-1 bg-yellow-100 text-yellow-800 text-xs font-semibold rounded-md hover:bg-yellow-200">Edit</button>
                    <button data-id="${param.id}" class="delete-param-btn px-3 py-1 bg-red-100 text-red-800 text-xs font-semibold rounded-md hover:bg-red-200">Delete</button>
                </div>
            </div>
            <div class="mt-2 text-sm">${mappingsHtml}</div>
        `;
        listContainer.appendChild(div);
    });

    listContainer.querySelectorAll('.duplicate-param-btn').forEach(btn => {
        btn.onclick = () => duplicateSpecialParameter(btn.dataset.id);
    });
    listContainer.querySelectorAll('.edit-param-btn').forEach(btn => {
        btn.onclick = () => editSpecialParameter(btn.dataset.id);
    });
    listContainer.querySelectorAll('.delete-param-btn').forEach(btn => {
        btn.onclick = () => deleteSpecialParameter(btn.dataset.id);
    });
}

function editSpecialParameter(paramId) {
    const param = specialParameters.find(p => p.id === paramId);
    if (param) {
        document.getElementById('manage-params-modal').classList.add('hidden');
        showSpecialParamModal(param);
    }
}

async function duplicateSpecialParameter(paramId) {
    let params = await getCustomParameters();
    const paramToDuplicate = params.find(p => p.id === paramId);

    if (!paramToDuplicate) {
        console.error("Parameter to duplicate not found");
        return;
    }

    const newParam = JSON.parse(JSON.stringify(paramToDuplicate)); // Deep copy

    // Find a unique name for the copy
    let newName = `${newParam.name}Copy`;
    const allParamNames = [...standardParameters, ...params.map(p => p.name)];
    while (allParamNames.includes(newName)) {
        newName = `${newName}Copy`;
    }
    
    newParam.name = newName;
    newParam.id = 'sparam_' + new Date().getTime(); // Assign new unique ID

    params.push(newParam);
    
    await saveCustomParameters(params);
    await loadCustomParameters(); 
    await populateParameterButtons();
    await showManageParamsModal(); // Refresh the list to show the new copy
}

async function deleteSpecialParameter(paramId) {
    let params = await getCustomParameters();
    params = params.filter(p => p.id !== paramId);
    await saveCustomParameters(params);
    await loadCustomParameters();
    await populateParameterButtons();
    await showManageParamsModal();
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
