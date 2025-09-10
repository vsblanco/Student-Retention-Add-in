import { DOM_IDS } from './constants.js';
import { getState } from './state.js';

let quill; // To hold the Quill editor instance

// --- Initialization ---

export function initializeQuill() {
    quill = new Quill('#' + DOM_IDS.EDITOR_CONTAINER, {
        theme: 'snow',
        modules: {
            toolbar: [
                ['bold', 'italic', 'underline'],
                [{ 'list': 'ordered'}, { 'list': 'bullet' }],
                ['link']
            ]
        }
    });
}

// --- View Management ---

export function showView(viewName) {
    const setupWizard = document.getElementById(DOM_IDS.SETUP_WIZARD);
    const emailComposer = document.getElementById(DOM_IDS.EMAIL_COMPOSER);

    if (viewName === 'composer') {
        setupWizard.classList.add('hidden');
        emailComposer.classList.remove('hidden');
    } else {
        setupWizard.classList.remove('hidden');
        emailComposer.classList.add('hidden');
    }
}

// --- UI Updates & Getters ---

export function updateStatus(message, color = 'gray', isSetupStatus = false) {
    const statusElement = document.getElementById(isSetupStatus ? DOM_IDS.SETUP_STATUS : DOM_IDS.STATUS);
    statusElement.textContent = message;
    statusElement.style.color = color;
}

export function getPowerAutomateUrl() {
    return document.getElementById(DOM_IDS.POWER_AUTOMATE_URL).value.trim();
}

export function getEmailTemplateFromDOM() {
    return {
        from: getPillboxValue(DOM_IDS.EMAIL_FROM_CONTAINER),
        recipientList: document.getElementById(DOM_IDS.RECIPIENT_LIST).value,
        customSheetName: document.getElementById(DOM_IDS.CUSTOM_SHEET_NAME).value.trim(),
        subject: getPillboxValue(DOM_IDS.EMAIL_SUBJECT_CONTAINER),
        cc: getPillboxValue(DOM_IDS.EMAIL_CC_CONTAINER),
        body: quill.root.innerHTML
    };
}


export function populateParameterButtons() {
    const { customParameters } = getState();
    const standardContainer = document.getElementById(DOM_IDS.STANDARD_PARAMETER_BUTTONS);
    const customContainer = document.getElementById(DOM_IDS.CUSTOM_PARAMETER_BUTTONS);
    const customSection = document.getElementById(DOM_IDS.CUSTOM_PARAMETERS_SECTION);

    // Standard parameters (hardcoded for now)
    const standardParams = ['StudentName', 'StudentEmail', 'Grade', 'LastLDA', 'DaysOut'];
    standardContainer.innerHTML = standardParams.map(p => 
        `<button class="parameter-button px-2 py-1 bg-indigo-100 text-indigo-800 text-xs rounded hover:bg-indigo-200" data-param="${p}">{${p}}</button>`
    ).join('');

    // Custom parameters from state
    if (customParameters && customParameters.length > 0) {
        customContainer.innerHTML = customParameters.map(p =>
            `<button class="parameter-button px-2 py-1 bg-blue-100 text-blue-800 text-xs rounded hover:bg-blue-200" data-param="${p.name}">{${p.name}}</button>`
        ).join('');
        customSection.classList.remove('hidden');
    } else {
        customSection.classList.add('hidden');
    }

    // Add event listeners to all parameter buttons
    document.querySelectorAll('.parameter-button').forEach(button => {
        button.onclick = () => insertTextInQuill(`{${button.dataset.param}}`);
    });
}


export function insertTextInQuill(text) {
    const range = quill.getSelection(true);
    quill.insertText(range.index, text, 'user');
    quill.setSelection(range.index + text.length, 0);
}


export function toggleCustomSheetInput() {
    const recipientList = document.getElementById(DOM_IDS.RECIPIENT_LIST);
    const customSheetContainer = document.getElementById(DOM_IDS.CUSTOM_SHEET_CONTAINER);
    if (recipientList.value === 'custom') {
        customSheetContainer.classList.remove('hidden');
    } else {
        customSheetContainer.classList.add('hidden');
    }
}

export function getSelectedSheetName() {
    const recipientList = document.getElementById(DOM_IDS.RECIPIENT_LIST);
    if (recipientList.value === 'custom') {
        return document.getElementById(DOM_IDS.CUSTOM_SHEET_NAME).value.trim() || 'Master List';
    }
    if (recipientList.value === 'lda') {
        const today = new Date();
        const month = String(today.getMonth() + 1).padStart(2, '0');
        const day = String(today.getDate()).padStart(2, '0');
        const year = today.getFullYear();
        return `${month}-${day}-${year}`;
    }
    return 'Master List';
}


// --- Modal UI ---

export function populateExampleModal(example) {
    document.getElementById(DOM_IDS.EXAMPLE_FROM).textContent = example.from || '';
    document.getElementById(DOM_IDS.EXAMPLE_TO).textContent = example.to || '';
    document.getElementById(DOM_IDS.EXAMPLE_CC).textContent = example.cc || '';
    document.getElementById(DOM_IDS.EXAMPLE_SUBJECT).textContent = example.subject || '';
    document.getElementById(DOM_IDS.EXAMPLE_BODY).innerHTML = example.body || '';
}

export function populatePayloadModal(payload) {
    const payloadContent = document.getElementById(DOM_IDS.PAYLOAD_CONTENT);
    const schemaContent = document.getElementById(DOM_IDS.SCHEMA_CONTENT);
    payloadContent.textContent = JSON.stringify(payload, null, 2);
    // You might want to generate a more formal JSON schema based on your data structure
    const schema = {
      "type": "array",
      "items": {
        "type": "object",
        "properties": {
          "to": { "type": "string", "format": "email" },
          "from": { "type": "string", "format": "email" },
          "cc": { "type": "string", "format": "email" },
          "subject": { "type": "string" },
          "body": { "type": "string", "contentMediaType": "text/html" }
        },
        "required": ["to", "from", "subject", "body"]
      }
    };
    schemaContent.textContent = JSON.stringify(schema, null, 2);
}

export function togglePayloadSchemaView() {
    const payloadContent = document.getElementById(DOM_IDS.PAYLOAD_CONTENT);
    const schemaContent = document.getElementById(DOM_IDS.SCHEMA_CONTENT);
    const button = document.getElementById(DOM_IDS.TOGGLE_PAYLOAD_SCHEMA_BUTTON);
    const title = document.getElementById(DOM_IDS.PAYLOAD_MODAL_TITLE);

    if (payloadContent.classList.contains('hidden')) {
        payloadContent.classList.remove('hidden');
        schemaContent.classList.add('hidden');
        button.textContent = 'Show Schema';
        title.textContent = 'Request Payload';
    } else {
        payloadContent.classList.add('hidden');
        schemaContent.classList.remove('hidden');
        button.textContent = 'Show Payload';
        title.textContent = 'Expected JSON Schema';
    }
}

export function populateSendConfirmationModal(count) {
    const message = document.getElementById(DOM_IDS.SEND_CONFIRM_MESSAGE);
    message.textContent = `You are about to send ${count} email(s). Do you want to proceed?`;
}


// --- Template UI ---

export function populateTemplatesModal(templates, handlers) {
    const container = document.getElementById(DOM_IDS.TEMPLATES_LIST_CONTAINER);
    if (!templates || templates.length === 0) {
        container.innerHTML = '<p class="text-center text-sm text-gray-500">No saved templates found.</p>';
        return;
    }
    container.innerHTML = templates.map(t => `
        <div class="border p-3 rounded-lg bg-gray-50 flex justify-between items-center">
            <div>
                <p class="font-semibold text-gray-800">${t.name}</p>
                <p class="text-xs text-gray-500">by ${t.author} on ${new Date(t.timestamp).toLocaleDateString()}</p>
            </div>
            <div class="space-x-2">
                <button class="load-template-btn px-3 py-1 bg-blue-100 text-blue-800 text-xs font-semibold rounded-md hover:bg-blue-200" data-id="${t.id}">Load</button>
                <button class="delete-template-btn px-3 py-1 bg-red-100 text-red-800 text-xs font-semibold rounded-md hover:bg-red-200" data-id="${t.id}">Delete</button>
            </div>
        </div>
    `).join('');

    container.querySelectorAll('.load-template-btn').forEach(btn => btn.onclick = () => handlers.onLoad(btn.dataset.id));
    container.querySelectorAll('.delete-template-btn').forEach(btn => btn.onclick = () => handlers.onDelete(btn.dataset.id));
}

export function getTemplateSaveForm() {
    return {
        name: document.getElementById(DOM_IDS.TEMPLATE_NAME).value.trim(),
        author: document.getElementById(DOM_IDS.TEMPLATE_AUTHOR).value.trim()
    };
}

export function clearSaveTemplateForm() {
    document.getElementById(DOM_IDS.TEMPLATE_NAME).value = '';
    document.getElementById(DOM_IDS.TEMPLATE_AUTHOR).value = '';
    updateSaveTemplateStatus('');
}

export function updateSaveTemplateStatus(message, color = 'gray') {
    const status = document.getElementById(DOM_IDS.SAVE_TEMPLATE_STATUS);
    status.textContent = message;
    status.style.color = color;
}

export function loadTemplateIntoForm(template) {
    // This is a simplified version. A real implementation would parse the pillbox values.
    setPillboxValue(DOM_IDS.EMAIL_FROM_CONTAINER, template.from || '');
    setPillboxValue(DOM_IDS.EMAIL_SUBJECT_CONTAINER, template.subject || '');
    setPillboxValue(DOM_IDS.EMAIL_CC_CONTAINER, template.cc || '');
    
    document.getElementById(DOM_IDS.RECIPIENT_LIST).value = template.recipientList || 'lda';
    document.getElementById(DOM_IDS.CUSTOM_SHEET_NAME).value = template.customSheetName || '';
    quill.root.innerHTML = template.body || '';
    toggleCustomSheetInput();
}


// --- Custom Parameter UI ---
export function showCustomParamModal(param = null) {
    const modal = document.getElementById(DOM_IDS.CUSTOM_PARAM_MODAL);
    const title = document.getElementById(DOM_IDS.CUSTOM_PARAM_MODAL_TITLE);
    
    // Clear form
    document.getElementById(DOM_IDS.PARAM_EDIT_ID).value = '';
    document.getElementById(DOM_IDS.PARAM_NAME).value = '';
    document.getElementById(DOM_IDS.PARAM_SOURCE_COLUMN).value = '';
    document.getElementById(DOM_IDS.PARAM_DEFAULT_VALUE).value = '';
    document.getElementById(DOM_IDS.PARAM_MAPPING_CONTAINER).innerHTML = '';
    updateCustomParamStatus('');
    
    if (param) {
        title.textContent = 'Edit Custom Parameter';
        document.getElementById(DOM_IDS.PARAM_EDIT_ID).value = param.id;
        document.getElementById(DOM_IDS.PARAM_NAME).value = param.name;
        document.getElementById(DOM_IDS.PARAM_SOURCE_COLUMN).value = param.sourceColumn;
        document.getElementById(DOM_IDS.PARAM_DEFAULT_VALUE).value = param.defaultValue;
        if (param.valueMap) {
            param.valueMap.forEach(mapping => addMappingRow(mapping.if, mapping.then));
        }
    } else {
        title.textContent = 'Create Custom Parameter';
        addMappingRow(); // Start with one empty mapping row
    }
    
    modal.classList.remove('hidden');
}

export function addMappingRow(ifValue = '', thenValue = '') {
    const container = document.getElementById(DOM_IDS.PARAM_MAPPING_CONTAINER);
    const row = document.createElement('div');
    row.className = 'flex items-center gap-2 text-sm';
    row.innerHTML = `
        <span class="font-semibold text-gray-500">If value is</span>
        <input type="text" class="mapping-if-input w-full px-2 py-1 border rounded-md" placeholder="e.g., Blanco, Victor" value="${ifValue}">
        <span class="font-semibold text-gray-500">then use</span>
        <input type="text" class="mapping-then-input w-full px-2 py-1 border rounded-md" placeholder="e.g., vblanco@school.edu" value="${thenValue}">
        <button type="button" class="remove-mapping-btn text-red-500 hover:text-red-700 font-bold text-lg">&times;</button>
    `;
    row.querySelector('.remove-mapping-btn').onclick = () => row.remove();
    container.appendChild(row);
}

export function getCustomParamForm() {
    const valueMap = [];
    document.querySelectorAll('#' + DOM_IDS.PARAM_MAPPING_CONTAINER + ' .flex').forEach(row => {
        const ifVal = row.querySelector('.mapping-if-input').value.trim();
        const thenVal = row.querySelector('.mapping-then-input').value.trim();
        if (ifVal && thenVal) {
            valueMap.push({ if: ifVal, then: thenVal });
        }
    });

    return {
        id: document.getElementById(DOM_IDS.PARAM_EDIT_ID).value,
        name: document.getElementById(DOM_IDS.PARAM_NAME).value.trim(),
        sourceColumn: document.getElementById(DOM_IDS.PARAM_SOURCE_COLUMN).value.trim(),
        defaultValue: document.getElementById(DOM_IDS.PARAM_DEFAULT_VALUE).value.trim(),
        valueMap: valueMap
    };
}

export function updateCustomParamStatus(message, color = 'gray') {
     const status = document.getElementById(DOM_IDS.SAVE_PARAM_STATUS);
     status.textContent = message;
     status.style.color = color;
}

export function populateManageCustomParamsModal(params, handlers) {
    const container = document.getElementById(DOM_IDS.MANAGE_PARAMS_LIST);
    if (!params || params.length === 0) {
        container.innerHTML = '<p class="text-sm text-gray-500 text-center">No custom parameters created yet.</p>';
        return;
    }
    container.innerHTML = params.map(p => `
        <div class="border p-3 rounded-lg bg-gray-50 flex justify-between items-start">
            <div>
                <p class="font-semibold text-gray-800">{${p.name}}</p>
                <p class="text-xs text-gray-500 mt-1">Source: <span class="font-medium">${p.sourceColumn}</span></p>
                ${p.defaultValue ? `<p class="text-xs text-gray-500">Default: <span class="font-medium">${p.defaultValue}</span></p>` : ''}
                ${p.valueMap && p.valueMap.length > 0 ? `<p class="text-xs text-gray-500">${p.valueMap.length} mapping rule(s)</p>` : ''}
            </div>
            <div class="flex-shrink-0 space-x-2">
                <button class="edit-param-btn px-3 py-1 bg-blue-100 text-blue-800 text-xs font-semibold rounded-md hover:bg-blue-200" data-id="${p.id}">Edit</button>
                <button class="duplicate-param-btn px-3 py-1 bg-yellow-100 text-yellow-800 text-xs font-semibold rounded-md hover:bg-yellow-200" data-id="${p.id}">Duplicate</button>
                <button class="delete-param-btn px-3 py-1 bg-red-100 text-red-800 text-xs font-semibold rounded-md hover:bg-red-200" data-id="${p.id}">Delete</button>
            </div>
        </div>
    `).join('');
    
    container.querySelectorAll('.edit-param-btn').forEach(btn => btn.onclick = () => handlers.onEdit(btn.dataset.id));
    container.querySelectorAll('.delete-param-btn').forEach(btn => btn.onclick = () => handlers.onDelete(btn.dataset.id));
    container.querySelectorAll('.duplicate-param-btn').forEach(btn => btn.onclick = () => handlers.onDuplicate(btn.dataset.id));
}

// --- Pillbox Input UI ---

export function setupPillboxInputs() {
    document.querySelectorAll('.pill-container').forEach(container => {
        const input = container.querySelector('.pill-input');
        
        container.addEventListener('click', () => input.focus());

        input.addEventListener('keydown', (e) => {
            if (e.key === 'Enter' || e.key === ',') {
                e.preventDefault();
                const value = input.value.trim();
                if (value) {
                    addPill(value, container);
                    input.value = '';
                }
            } else if (e.key === 'Backspace' && input.value === '') {
                const lastPill = container.querySelector('.pill-tag:last-of-type');
                if (lastPill) {
                    input.value = lastPill.textContent.slice(0, -1); // Put text back in input, remove 'x'
                    lastPill.remove();
                }
            }
        });

        input.addEventListener('blur', () => {
             const value = input.value.trim();
             if (value) {
                 addPill(value, container);
                 input.value = '';
             }
        });
    });
}

function addPill(text, container) {
    const pill = document.createElement('span');
    pill.className = 'pill-tag';
    
    // Check if it's a parameter
    if (text.startsWith('{') && text.endsWith('}')) {
        pill.classList.add('param');
    }

    const textNode = document.createTextNode(text);
    const removeBtn = document.createElement('span');
    removeBtn.className = 'pill-remove';
    removeBtn.innerHTML = '&times;';
    removeBtn.onclick = (e) => {
        e.stopPropagation(); // prevent container click from firing
        pill.remove();
    };
    
    pill.appendChild(textNode);
    pill.appendChild(removeBtn);
    
    const input = container.querySelector('.pill-input');
    container.insertBefore(pill, input);
}

function getPillboxValue(containerId) {
    const container = document.getElementById(containerId);
    const pills = Array.from(container.querySelectorAll('.pill-tag'));
    return pills.map(p => p.textContent.slice(0, -1)).join(', '); // remove the 'x'
}

function setPillboxValue(containerId, valueString) {
    const container = document.getElementById(containerId);
    // Clear existing pills
    container.querySelectorAll('.pill-tag').forEach(p => p.remove());
    if (valueString) {
        valueString.split(',').forEach(val => {
            const trimmedVal = val.trim();
            if (trimmedVal) {
                addPill(trimmedVal, container);
            }
        });
    }
}


// --- Randomize Modal UI ---

export function clearRandomizeModal() {
    const container = document.getElementById(DOM_IDS.RANDOMIZE_OPTIONS_CONTAINER);
    container.innerHTML = '';
    // Add two empty options to start
    addRandomizeOption();
    addRandomizeOption();
}

export function addRandomizeOption(value = '') {
    const container = document.getElementById(DOM_IDS.RANDOMIZE_OPTIONS_CONTAINER);
    const optionDiv = document.createElement('div');
    optionDiv.className = 'flex items-center gap-2';
    
    optionDiv.innerHTML = `
        <input type="text" class="randomize-option-input w-full px-2 py-1 border rounded-md" placeholder="Enter a text variation..." value="${value}">
        <button type="button" class="remove-randomize-option-btn text-red-500 hover:text-red-700 font-bold text-lg">&times;</button>
    `;
    
    optionDiv.querySelector('.remove-randomize-option-btn').onclick = () => {
        // Prevent removing the last two options
        if (container.childElementCount > 2) {
            optionDiv.remove();
        }
    };
    
    container.appendChild(optionDiv);
}

export function getRandomizeOptions() {
    const inputs = document.querySelectorAll('#' + DOM_IDS.RANDOMIZE_OPTIONS_CONTAINER + ' .randomize-option-input');
    return Array.from(inputs).map(input => input.value.trim()).filter(val => val);
}
