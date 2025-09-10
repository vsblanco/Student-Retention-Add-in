import { DOM_IDS } from './constants.js';
import { updateState, getState, setQuill, getQuill } from './state.js';

let RandomizeBlot, ConditionBlot;

// --- Initialization ---

export function initializeQuill() {
    try {
        const quill = new Quill(`#${DOM_IDS.EDITOR_CONTAINER}`, {
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
        console.log("[LOG] Quill initialized successfully.");
        setQuill(quill); // Use the dedicated setter for the Quill instance.
        registerCustomBlots(Quill);

        quill.on('selection-change', (range) => {
            if (range) {
                updateState('lastFocusedElement', quill);
            }
        });
    } catch (error) {
        console.error("[ERROR] Quill initialization failed:", error);
    }
}

/**
 * Registers the custom Blot formats with Quill.
 * @param {object} Quill The Quill instance.
 */
function registerCustomBlots(Quill) {
    const Inline = Quill.import('blots/inline');

    // --- ParameterBlot ---
    class ParameterBlot extends Inline {
        static create(value) {
            let node = super.create();
            node.setAttribute('contenteditable', 'false');
            node.setAttribute('data-param', value.name);
            node.classList.add('parameter-tag');
            if (value.isCustom) node.classList.add('custom');
            node.innerText = `{${value.name}}`;
            return node;
        }
        static formats(node) {
            return { name: node.getAttribute('data-param'), isCustom: node.classList.contains('custom') };
        }
    }
    ParameterBlot.blotName = 'parameter';
    ParameterBlot.tagName = 'SPAN';
    Quill.register(ParameterBlot);
    
    // --- RandomizeBlot (Simplified) ---
    RandomizeBlot = class extends Inline {
        static create(value) {
            const node = super.create(value);
            node.setAttribute('contenteditable', 'false');
            node.classList.add('randomize-tag');
            node.innerText = `{Randomize}`;
            return node;
        }
        static value(node) {
            return true;
        }
    }
    RandomizeBlot.blotName = 'randomize';
    RandomizeBlot.tagName = 'SPAN';
    Quill.register(RandomizeBlot);

    // --- ConditionBlot (Simplified) ---
    ConditionBlot = class extends Inline {
        static create(value) {
            const node = super.create(value);
            node.setAttribute('contenteditable', 'false');
            node.classList.add('condition-tag');
            node.innerText = `{Condition}`;
            return node;
        }
        static value(node) {
            return true;
        }
    }
    ConditionBlot.blotName = 'condition';
    ConditionBlot.tagName = 'SPAN';
    Quill.register(ConditionBlot);
}


// --- DOM Manipulation & UI Logic ---

export function showView(viewName) {
    const setupWizard = document.getElementById(DOM_IDS.SETUP_WIZARD);
    const emailComposer = document.getElementById(DOM_IDS.EMAIL_COMPOSER);

    if (viewName === 'composer') {
        setupWizard.classList.add('hidden');
        emailComposer.classList.remove('hidden');
    } else { // 'setup'
        setupWizard.classList.remove('hidden');
        emailComposer.classList.add('hidden');
    }
}

export function updateStatus(message, color = 'gray', isSetupStatus = false) {
    const statusEl = document.getElementById(isSetupStatus ? DOM_IDS.SETUP_STATUS : DOM_IDS.STATUS);
    if (!statusEl) return;
    statusEl.textContent = message;
    statusEl.style.color = color;
}

export function getPowerAutomateUrl() {
    return document.getElementById(DOM_IDS.POWER_AUTOMATE_URL).value.trim();
}

export function getSelectedSheetName() {
    const recipientList = document.getElementById(DOM_IDS.RECIPIENT_LIST);
    if (recipientList.value === 'custom') {
        return document.getElementById(DOM_IDS.CUSTOM_SHEET_NAME).value.trim();
    }
    return recipientList.value === 'lda' ? getTodaysLdaSheetName() : 'Master List';
}

function getTodaysLdaSheetName() {
    const now = new Date();
    return `LDA ${now.getMonth() + 1}-${now.getDate()}-${now.getFullYear()}`;
}

export function toggleCustomSheetInput() {
    const recipientList = document.getElementById(DOM_IDS.RECIPIENT_LIST);
    const customSheetContainer = document.getElementById(DOM_IDS.CUSTOM_SHEET_CONTAINER);
    customSheetContainer.classList.toggle('hidden', recipientList.value !== 'custom');
}


// --- Parameter Button & Insertion Logic ---

export function populateParameterButtons() {
    const { customParameters } = getState();
    const standardContainer = document.getElementById(DOM_IDS.STANDARD_PARAMETER_BUTTONS);
    const customContainer = document.getElementById(DOM_IDS.CUSTOM_PARAMETER_BUTTONS);
    const customSection = document.getElementById(DOM_IDS.CUSTOM_PARAMETERS_SECTION);

    standardContainer.innerHTML = '';
    customContainer.innerHTML = '';

    const createButton = (param, container, isCustom = false) => {
        const button = document.createElement('button');
        button.className = isCustom ? 'px-2 py-1 bg-blue-100 text-blue-800 text-xs rounded hover:bg-blue-200' : 'px-2 py-1 bg-gray-200 text-gray-800 text-xs rounded hover:bg-gray-300';
        button.textContent = `{${param}}`;
        button.onmousedown = (e) => {
            e.preventDefault(); 
            insertParameter(param, isCustom);
        };
        container.appendChild(button);
    };

    const standardParameters = ['FirstName', 'LastName', 'StudentName', 'StudentEmail', 'PersonalEmail', 'Grade', 'DaysOut'];
    standardParameters.forEach(param => createButton(param, standardContainer, false));

    if (customParameters.length > 0) {
        customSection.classList.remove('hidden');
        customParameters.forEach(param => createButton(param.name, customContainer, true));
    } else {
        customSection.classList.add('hidden');
    }

    document.getElementById(DOM_IDS.RANDOMIZE_PARAMETER_BUTTON).onmousedown = (e) => { e.preventDefault(); insertSpecialParameter('randomize'); };
    document.getElementById(DOM_IDS.CONDITION_PARAMETER_BUTTON).onmousedown = (e) => { e.preventDefault(); insertSpecialParameter('condition'); };
}

function insertParameter(paramName, isCustom) {
    const { lastFocusedElement } = getState();
    const quill = getQuill();
    
    if (!quill) {
        console.error("[ERROR] insertParameter called but Quill instance is not available.");
        return;
    }

    if (lastFocusedElement && lastFocusedElement.classList && lastFocusedElement.classList.contains('param-input')) {
        lastFocusedElement.value = `{${paramName}}`;
        lastFocusedElement.dispatchEvent(new Event('input', { bubbles: true }));
    } else if (lastFocusedElement && ['email-from-input', 'email-subject-input', 'email-cc-input'].includes(lastFocusedElement.id)) {
        const type = lastFocusedElement.id.replace('email-', '').replace('-input', '');
        addPill(type, `{${paramName}}`);
    } else {
        const range = quill.getSelection(true);
        quill.insertEmbed(range.index, 'parameter', { name: paramName, isCustom }, Quill.sources.USER);
        // FIX: Use a timeout to ensure the cursor is placed correctly after the browser has rendered the blot.
        setTimeout(() => {
            quill.setSelection(range.index + 1, Quill.sources.SILENT);
        }, 0);
    }
}

function insertSpecialParameter(type) {
    const quill = getQuill();

    if (!quill) {
        console.error("[ERROR] insertSpecialParameter called but Quill instance is not available.");
        return;
    }

    const range = quill.getSelection(true);
    // Pass a simple `true` value, as the blot is no longer complex.
    quill.insertEmbed(range.index, type, true, Quill.sources.USER);
    // FIX: Use a timeout to ensure the cursor is placed correctly after the browser has rendered the blot.
    setTimeout(() => {
        quill.setSelection(range.index + 1, Quill.sources.SILENT);
    }, 0);
}


// --- Pillbox Input UI ---

export function setupPillboxInputs() {
    ['from', 'subject', 'cc'].forEach(type => {
        const container = document.getElementById(`email-${type}-container`);
        const input = document.getElementById(`email-${type}-input`);

        container.addEventListener('click', () => {
            updateState('lastFocusedElement', input);
            input.focus();
        });

        input.addEventListener('focus', () => updateState('lastFocusedElement', input));

        input.addEventListener('keydown', (e) => {
            if (e.key === ',' || e.key === 'Enter' || e.key === ';') {
                e.preventDefault();
                addPill(type, input.value.trim());
                input.value = '';
            } else if (e.key === 'Backspace' && input.value === '') {
                removeLastPill(type);
            }
        });

        input.addEventListener('blur', () => {
            addPill(type, input.value.trim());
            input.value = '';
        });
    });
}

function addPill(type, text) {
    if (!text) return;
    const { emailParts } = getState();
    emailParts[type].push(text);
    updateState('emailParts', emailParts);
    renderPills(type);
}

function removeLastPill(type) {
    const { emailParts } = getState();
    if (emailParts[type].length > 0) {
        emailParts[type].pop();
        updateState('emailParts', emailParts);
        renderPills(type);
    }
}

function removePill(type, index) {
    const { emailParts } = getState();
    emailParts[type].splice(index, 1);
    updateState('emailParts', emailParts);
    renderPills(type);
}

function renderPills(type) {
    const { emailParts } = getState();
    const container = document.getElementById(`email-${type}-container`);
    const input = document.getElementById(`email-${type}-input`);
    
    container.querySelectorAll('.pill-tag').forEach(pill => pill.remove());

    emailParts[type].forEach((part, index) => {
        const isParam = part.startsWith('{') && part.endsWith('}');
        const pill = document.createElement('span');
        pill.className = isParam ? 'pill-tag param' : 'pill-tag';
        pill.textContent = part;
        
        const removeBtn = document.createElement('span');
        removeBtn.textContent = '×';
        removeBtn.className = 'pill-remove';
        removeBtn.onclick = (e) => {
            e.stopPropagation();
            removePill(type, index);
        };
        
        pill.appendChild(removeBtn);
        container.insertBefore(pill, input);
    });
}

function reconstructPillboxString(parts, separator = '') {
    return parts.join(separator);
}

export function getEmailTemplateFromDOM() {
    const { emailParts } = getState();
    const quill = getQuill();

    if (!quill) {
        console.error("[ERROR] getEmailTemplateFromDOM called but Quill instance is not available.");
        return {
            from: reconstructPillboxString(emailParts.from),
            subject: reconstructPillboxString(emailParts.subject),
            cc: reconstructPillboxString(emailParts.cc, ';'),
            body: '<p></p>'
        };
    }

    const editor = quill.root;
    
    // This logic is no longer needed with the simplified blots.
    // editor.querySelectorAll('.randomize-tag, .condition-tag').forEach(tagNode => {
    //     const blot = Quill.find(tagNode, true);
    //     if (blot && blot.statics.blotName === 'randomize') {
    //         RandomizeBlot.updateOptions(tagNode);
    //     } else if (blot && blot.statics.blotName === 'condition') {
    //         ConditionBlot.updateOptions(tagNode);
    //     }
    // });

    return {
        from: reconstructPillboxString(emailParts.from),
        subject: reconstructPillboxString(emailParts.subject),
        cc: reconstructPillboxString(emailParts.cc, ';'),
        body: editor.innerHTML
    };
}

export function loadTemplateIntoForm(template) {
    const parsePills = (str) => {
        if (!str) return [];
        return str.split(/({[^}]+})/g).filter(part => part);
    };
    
    const newEmailParts = {
        from: parsePills(template.from || ''),
        subject: parsePills(template.subject || ''),
        cc: parsePills(template.cc || '')
    };
    updateState('emailParts', newEmailParts);

    renderPills('from');
    renderPills('subject');
    renderPills('cc');
    
    const quill = getQuill();
    if (quill) {
        quill.root.innerHTML = template.body || '<p></p>';
    }
}


// --- Modal Content Management ---

export function populateExampleModal(examplePayload) {
    document.getElementById(DOM_IDS.EXAMPLE_FROM).textContent = examplePayload.from || '[Not Specified]';
    document.getElementById(DOM_IDS.EXAMPLE_TO).textContent = examplePayload.to || '[No Email Found]';
    document.getElementById(DOM_IDS.EXAMPLE_CC).textContent = examplePayload.cc || '[Not Specified]';
    document.getElementById(DOM_IDS.EXAMPLE_SUBJECT).textContent = examplePayload.subject;
    document.getElementById(DOM_IDS.EXAMPLE_BODY).innerHTML = examplePayload.body;
}

export function populatePayloadModal(payload) {
    document.getElementById(DOM_IDS.PAYLOAD_CONTENT).textContent = JSON.stringify(payload, null, 2);
    document.getElementById(DOM_IDS.PAYLOAD_CONTENT).classList.remove('hidden');
    document.getElementById(DOM_IDS.SCHEMA_CONTENT).classList.add('hidden');
    document.getElementById(DOM_IDS.PAYLOAD_MODAL_TITLE).textContent = 'Request Payload';
    document.getElementById(DOM_IDS.TOGGLE_PAYLOAD_SCHEMA_BUTTON).textContent = 'Show Schema';
}

export function togglePayloadSchemaView() {
    const payloadContent = document.getElementById(DOM_IDS.PAYLOAD_CONTENT);
    const schemaContent = document.getElementById(DOM_IDS.SCHEMA_CONTENT);
    const title = document.getElementById(DOM_IDS.PAYLOAD_MODAL_TITLE);
    const button = document.getElementById(DOM_IDS.TOGGLE_PAYLOAD_SCHEMA_BUTTON);

    if (payloadContent.classList.contains('hidden')) {
        payloadContent.classList.remove('hidden');
        schemaContent.classList.add('hidden');
        title.textContent = 'Request Payload';
        button.textContent = 'Show Schema';
    } else {
        payloadContent.classList.add('hidden');
        schemaContent.classList.remove('hidden');
        title.textContent = 'Request Body JSON Schema';
        button.textContent = 'Show Payload';
    }
}

export function populateSendConfirmationModal(count) {
    document.getElementById(DOM_IDS.SEND_CONFIRM_MESSAGE).textContent = `You are about to send emails to ${count} student(s). Do you want to proceed?`;
}

export function populateTemplatesModal(templates, handlers) {
    const container = document.getElementById(DOM_IDS.TEMPLATES_LIST_CONTAINER);
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
            </div>`;
        container.appendChild(div);
    });
    
    container.querySelectorAll('.load-template-btn').forEach(btn => btn.onclick = () => handlers.onLoad(btn.dataset.id));
    container.querySelectorAll('.delete-template-btn').forEach(btn => btn.onclick = () => handlers.onDelete(btn.dataset.id));
}

export function clearSaveTemplateForm() {
    document.getElementById(DOM_IDS.TEMPLATE_NAME).value = '';
    document.getElementById(DOM_IDS.TEMPLATE_AUTHOR).value = '';
    updateSaveTemplateStatus('');
}

export function getTemplateSaveForm() {
    return {
        name: document.getElementById(DOM_IDS.TEMPLATE_NAME).value.trim(),
        author: document.getElementById(DOM_IDS.TEMPLATE_AUTHOR).value.trim()
    };
}

export function updateSaveTemplateStatus(message, color = 'gray') {
    const statusEl = document.getElementById(DOM_IDS.SAVE_TEMPLATE_STATUS);
    statusEl.textContent = message;
    statusEl.style.color = color;
}

export function showCustomParamModal(paramToEdit = null) {
    const modalTitle = document.getElementById(DOM_IDS.CUSTOM_PARAM_MODAL_TITLE);
    const nameInput = document.getElementById(DOM_IDS.PARAM_NAME);
    const sourceColumnInput = document.getElementById(DOM_IDS.PARAM_SOURCE_COLUMN);
    const defaultValueInput = document.getElementById(DOM_IDS.PARAM_DEFAULT_VALUE);
    const mappingContainer = document.getElementById(DOM_IDS.PARAM_MAPPING_CONTAINER);
    const editIdInput = document.getElementById(DOM_IDS.PARAM_EDIT_ID);

    mappingContainer.innerHTML = '';
    updateCustomParamStatus('');

    if (paramToEdit) {
        modalTitle.textContent = 'Edit Custom Parameter';
        editIdInput.value = paramToEdit.id;
        nameInput.value = paramToEdit.name;
        sourceColumnInput.value = paramToEdit.sourceColumn;
        defaultValueInput.value = paramToEdit.defaultValue;
        (paramToEdit.mappings || []).forEach(m => addMappingRow(m.if, m.then));
    } else {
        modalTitle.textContent = 'Create Custom Parameter';
        editIdInput.value = '';
        nameInput.value = '';
        sourceColumnInput.value = '';
        defaultValueInput.value = '';
    }
    document.getElementById(DOM_IDS.CUSTOM_PARAM_MODAL).classList.remove('hidden');
}

export function getCustomParamForm() {
    const mappings = [];
    document.querySelectorAll(`#${DOM_IDS.PARAM_MAPPING_CONTAINER} .mapping-row`).forEach(row => {
        const ifValue = row.querySelector('.mapping-if').value.trim();
        const thenValue = row.querySelector('.mapping-then').value.trim();
        if (ifValue) {
            mappings.push({ if: ifValue, then: thenValue });
        }
    });

    return {
        id: document.getElementById(DOM_IDS.PARAM_EDIT_ID).value,
        name: document.getElementById(DOM_IDS.PARAM_NAME).value.trim(),
        sourceColumn: document.getElementById(DOM_IDS.PARAM_SOURCE_COLUMN).value.trim(),
        defaultValue: document.getElementById(DOM_IDS.PARAM_DEFAULT_VALUE).value.trim(),
        mappings: mappings
    };
}

export function addMappingRow(ifValue = '', thenValue = '') {
    const container = document.getElementById(DOM_IDS.PARAM_MAPPING_CONTAINER);
    const div = document.createElement('div');
    div.className = 'flex items-center gap-2 mapping-row';
    div.innerHTML = `
        <span class="text-sm text-gray-500">If cell is</span>
        <input type="text" class="mapping-if flex-1 px-2 py-1 border border-gray-300 rounded-md text-sm" placeholder="e.g., Bob" value="${ifValue}">
        <span class="text-sm text-gray-500">then value is</span>
        <input type="text" class="mapping-then flex-1 px-2 py-1 border border-gray-300 rounded-md text-sm" placeholder="e.g., bobjones@gmail.com" value="${thenValue}">
        <button class="remove-mapping-btn text-red-500 hover:text-red-700 text-lg">&times;</button>`;
    div.querySelector('.remove-mapping-btn').onclick = () => div.remove();
    container.appendChild(div);
}

export function updateCustomParamStatus(message, color = 'gray') {
    const statusEl = document.getElementById(DOM_IDS.SAVE_PARAM_STATUS);
    statusEl.textContent = message;
    statusEl.style.color = color;
}

export function populateManageCustomParamsModal(params, handlers) {
    const listContainer = document.getElementById(DOM_IDS.MANAGE_PARAMS_LIST);
    listContainer.innerHTML = '';
    if (params.length === 0) {
        listContainer.innerHTML = '<p class="text-gray-500 text-center">No custom parameters created yet.</p>';
        return;
    }

    params.forEach(param => {
        const div = document.createElement('div');
        div.className = 'p-3 border-b';
        let mappingsHtml = (param.mappings || []).map(m => `<div class="text-xs ml-4"><span class="text-gray-500">If '${m.if}' &rarr;</span> '${m.then}'</div>`).join('');
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
            <div class="mt-2 text-sm">${mappingsHtml}</div>`;
        listContainer.appendChild(div);
    });

    listContainer.querySelectorAll('.duplicate-param-btn').forEach(btn => btn.onclick = () => handlers.onDuplicate(btn.dataset.id));
    listContainer.querySelectorAll('.edit-param-btn').forEach(btn => btn.onclick = () => handlers.onEdit(btn.dataset.id));
    listContainer.querySelectorAll('.delete-param-btn').forEach(btn => btn.onclick = () => handlers.onDelete(btn.dataset.id));
}

