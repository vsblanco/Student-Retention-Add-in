import { DOM_IDS, STANDARD_PARAMETERS } from './constants.js';
import { getState, updateState, getQuill, setQuill } from './state.js';
import { getTodaysLdaSheetName } from './utils.js';


// A mutable variable to hold the Quill editor instance.
// It's kept separate from appState because it's a complex object, not simple data.
let quillInstance = null;

// Define Blot classes in a scope accessible by all functions
let RandomizeBlot;
let ConditionBlot;

export function initializeQuill() {
    // Define custom formats (Blots) for Quill
    const Embed = Quill.import('blots/embed');
    const Inline = Quill.import('blots/inline');

    class ParameterBlot extends Inline {
        static create(value) {
            const node = super.create();
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

    RandomizeBlot = class extends Embed {
        static create(value) {
            const wrapper = document.createElement('span');
            wrapper.classList.add('randomize-tag-wrapper');
            wrapper.setAttribute('contenteditable', 'false');

            const node = document.createElement('span');
            node.classList.add('randomize-tag');
            
            const text = document.createElement('span');
            text.innerText = '{Randomize}';
            
            const arrow = document.createElement('span');
            arrow.innerHTML = '&#9660;';
            arrow.classList.add('randomize-arrow');
            
            node.appendChild(text);
            node.appendChild(arrow);
            
            const panel = document.createElement('div');
            panel.classList.add('randomize-panel');
            panel.style.display = 'none';
            
            const inputsContainer = document.createElement('div');
            (value.options || ['']).forEach(optionText => {
                this.addOptionInput(inputsContainer, optionText, wrapper);
            });
            
            const addButton = document.createElement('button');
            addButton.innerText = '+ Add Option';
            addButton.classList.add('add-random-option');
            addButton.onclick = (e) => {
                e.stopPropagation();
                this.addOptionInput(inputsContainer, '', wrapper);
            };
            
            panel.appendChild(inputsContainer);
            panel.appendChild(addButton);
            
            wrapper.appendChild(node);
            wrapper.appendChild(panel);

            node.onclick = (e) => {
                e.stopPropagation();
                const isOpening = panel.style.display === 'none';
                document.querySelectorAll('.randomize-panel, .condition-panel').forEach(p => { if (p !== panel) p.style.display = 'none'; });
                document.querySelectorAll('.randomize-arrow, .condition-arrow').forEach(a => { if (a !== arrow) a.classList.remove('open'); });

                panel.style.display = isOpening ? 'block' : 'none';
                arrow.classList.toggle('open', isOpening);
            };

            panel.addEventListener('click', e => e.stopPropagation());

            this.updateOptions(wrapper);
            return wrapper;
        }

        static addOptionInput(container, value, blotNode) {
            const optionWrapper = document.createElement('div');
            optionWrapper.classList.add('randomize-option-wrapper');

            const textarea = document.createElement('textarea');
            textarea.classList.add('randomize-input');
            textarea.value = value;
            textarea.placeholder = 'Enter a phrase...';
            textarea.rows = 1;

            const autoResize = () => {
                textarea.style.height = 'auto';
                textarea.style.height = textarea.scrollHeight + 'px';
                this.updateOptions(blotNode);
            };
            
            textarea.oninput = autoResize;
            
            const deleteBtn = document.createElement('button');
            deleteBtn.innerHTML = '&times;';
            deleteBtn.classList.add('randomize-option-delete');
            deleteBtn.onclick = (e) => {
                e.stopPropagation();
                optionWrapper.remove();
                this.updateOptions(blotNode);
            };

            optionWrapper.appendChild(textarea);
            optionWrapper.appendChild(deleteBtn);
            container.appendChild(optionWrapper);

            setTimeout(autoResize, 0);
        }

        static updateOptions(blotNode) {
            const inputs = blotNode.querySelectorAll('.randomize-input');
            const options = Array.from(inputs).map(input => input.value);
            blotNode.dataset.options = JSON.stringify(options.filter(o => o.trim()));
        }

        static value(domNode) {
            return {
                options: JSON.parse(domNode.dataset.options || '[]')
            };
        }
    };
    RandomizeBlot.blotName = 'randomize';
    RandomizeBlot.tagName = 'SPAN';
    Quill.register(RandomizeBlot);

    ConditionBlot = class extends Embed {
        static create(value) {
            const wrapper = document.createElement('span');
            wrapper.classList.add('condition-tag-wrapper');
            wrapper.setAttribute('contenteditable', 'false');

            const node = document.createElement('span');
            node.classList.add('condition-tag');
            const text = document.createElement('span');
            text.innerText = '{Condition}';
            const arrow = document.createElement('span');
            arrow.innerHTML = '&#9660;';
            arrow.classList.add('condition-arrow');
            node.appendChild(text);
            node.appendChild(arrow);

            const panel = document.createElement('div');
            panel.classList.add('condition-panel');
            panel.style.display = 'none';

            const ifClause = document.createElement('div');
            ifClause.classList.add('condition-clause');
            const ifParamInput = document.createElement('input');
            ifParamInput.className = 'condition-input param-input';
            ifParamInput.placeholder = '{Parameter}';
            ifParamInput.value = value.if_param ? `{${value.if_param}}` : '';
            ifParamInput.onfocus = () => { updateState('lastFocusedElement', ifParamInput); };
            
            ifClause.innerHTML = `<span class="condition-keyword">IF</span>`;
            ifClause.appendChild(ifParamInput);
            ifClause.innerHTML += `
                <select class="condition-operator">
                    ${['=', '>', '>=', '<', '<='].map(op => `<option ${op === value.operator ? 'selected' : ''}>${op}</option>`).join('')}
                </select>
                <input class="condition-input value-input" placeholder="Value" value="${value.if_value || ''}">
            `;

            const thenClause = document.createElement('div');
            thenClause.classList.add('condition-clause');
            thenClause.innerHTML = `<span class="condition-keyword then">THEN</span>`;
            const thenTextarea = document.createElement('textarea');
            thenTextarea.classList.add('condition-then-input');
            thenTextarea.placeholder = 'Enter text to show if true... You can use {Parameters} here.';
            thenTextarea.rows = 2;
            thenTextarea.value = value.then_text || '';
            thenClause.appendChild(thenTextarea);

            panel.appendChild(ifClause);
            panel.appendChild(thenClause);
            wrapper.appendChild(node);
            wrapper.appendChild(panel);
            
            const update = () => this.updateOptions(wrapper);
            panel.querySelectorAll('input, select, textarea').forEach(el => el.oninput = update);

            node.onclick = (e) => {
                e.stopPropagation();
                const isOpening = panel.style.display === 'none';
                document.querySelectorAll('.randomize-panel, .condition-panel').forEach(p => { if (p !== panel) p.style.display = 'none'; });
                document.querySelectorAll('.randomize-arrow, .condition-arrow').forEach(a => { if (a !== arrow) a.classList.remove('open'); });

                panel.style.display = isOpening ? 'block' : 'none';
                arrow.classList.toggle('open', isOpening);
            };
            
            panel.addEventListener('click', e => e.stopPropagation());
            this.updateOptions(wrapper);
            return wrapper;
        }
        
        static updateOptions(blotNode) {
            const data = {
                if_param: blotNode.querySelector('.param-input').value.replace(/[{}]/g, ''),
                operator: blotNode.querySelector('.condition-operator').value,
                if_value: blotNode.querySelector('.value-input').value,
                then_text: blotNode.querySelector('.condition-then-input').value
            };
            blotNode.dataset.condition = JSON.stringify(data);
        }

        static value(domNode) {
            return JSON.parse(domNode.dataset.condition || '{}');
        }
    };
    ConditionBlot.blotName = 'condition';
    ConditionBlot.tagName = 'SPAN';
    Quill.register(ConditionBlot);

    // Initialize Quill
    const quill = new Quill(`#${DOM_IDS.EDITOR_CONTAINER}`, {
        theme: 'snow',
        modules: {
            toolbar: [
                ['bold', 'italic', 'underline'],
                [{ 'list': 'ordered' }, { 'list': 'bullet' }],
                [{ 'color': [] }, { 'background': [] }],
                ['link']
            ]
        }
    });

    setQuill(quill); // Save the instance to the state module

    quill.on('selection-change', (range) => {
        if (range) {
            updateState('lastFocusedElement', quill);
        }
    });
}

// --- View Management ---

export function showView(view) {
    const setupWizard = document.getElementById(DOM_IDS.SETUP_WIZARD);
    const emailComposer = document.getElementById(DOM_IDS.EMAIL_COMPOSER);
    if (view === 'composer') {
        setupWizard.classList.add('hidden');
        emailComposer.classList.remove('hidden');
    } else {
        setupWizard.classList.remove('hidden');
        emailComposer.classList.add('hidden');
    }
}

// --- Parameter UI ---

export function populateParameterButtons() {
    const { customParameters } = getState();
    const standardContainer = document.getElementById(DOM_IDS.STANDARD_PARAMETER_BUTTONS);
    const customContainer = document.getElementById(DOM_IDS.CUSTOM_PARAMETER_BUTTONS);
    const customSection = document.getElementById(DOM_IDS.CUSTOM_PARAMETERS_SECTION);

    standardContainer.innerHTML = '';
    customContainer.innerHTML = '';

    const createButton = (param, container, isCustom = false) => {
        const button = document.createElement('button');
        button.className = isCustom
            ? 'px-2 py-1 bg-blue-100 text-blue-800 text-xs rounded hover:bg-blue-200'
            : 'px-2 py-1 bg-gray-200 text-gray-800 text-xs rounded hover:bg-gray-300';
        button.textContent = `{${param}}`;
        button.onmousedown = (e) => {
            e.preventDefault();
            insertParameter(`{${param}}`);
        };
        container.appendChild(button);
    };

    STANDARD_PARAMETERS.forEach(param => createButton(param, standardContainer, false));

    if (customParameters.length > 0) {
        customSection.classList.remove('hidden');
        customParameters.forEach(param => createButton(param.name, customContainer, true));
    } else {
        customSection.classList.add('hidden');
    }

    document.getElementById(DOM_IDS.RANDOMIZE_PARAMETER_BUTTON).onmousedown = (e) => { e.preventDefault(); insertParameter('{Randomize}'); };
    document.getElementById(DOM_IDS.CONDITION_PARAMETER_BUTTON).onmousedown = (e) => { e.preventDefault(); insertParameter('{Condition}'); };
}

function insertParameter(param) {
    const { lastFocusedElement } = getState();
    const quill = getQuill();
    const paramName = param.replace(/[{}]/g, '');

    if (lastFocusedElement && lastFocusedElement.classList && lastFocusedElement.classList.contains('param-input')) {
        lastFocusedElement.value = param;
        lastFocusedElement.dispatchEvent(new Event('input', { bubbles: true }));
    } else if (lastFocusedElement && ['email-from-input', 'email-subject-input', 'email-cc-input'].includes(lastFocusedElement.id)) {
        const type = lastFocusedElement.id.replace('email-', '').replace('-input', '');
        addPill(type, param);
    } else {
        const range = quill.getSelection(true);
        if (paramName === 'Randomize') {
            quill.insertEmbed(range.index, 'randomize', { options: [''] }, Quill.sources.USER);
        } else if (paramName === 'Condition') {
            quill.insertEmbed(range.index, 'condition', {}, Quill.sources.USER);
        } else {
            const isCustom = getState().customParameters.some(p => p.name === paramName);
            quill.insertEmbed(range.index, 'parameter', { name: paramName, isCustom }, Quill.sources.USER);
        }
        quill.setSelection(range.index + 1, Quill.sources.USER);
    }
}


// --- Pillbox UI ---

export function setupPillboxInputs() {
    ['from', 'subject', 'cc'].forEach(type => {
        const container = document.getElementById(DOM_IDS[`${type.toUpperCase()}_CONTAINER`]);
        const input = document.getElementById(DOM_IDS[`${type.toUpperCase()}_INPUT`]);

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
    renderPills(type);
}

function removeLastPill(type) {
    const { emailParts } = getState();
    if (emailParts[type].length > 0) {
        emailParts[type].pop();
        renderPills(type);
    }
}

function removePill(type, index) {
    const { emailParts } = getState();
    emailParts[type].splice(index, 1);
    renderPills(type);
}

function renderPills(type) {
    const { emailParts } = getState();
    const container = document.getElementById(DOM_IDS[`${type.toUpperCase()}_CONTAINER`]);
    const input = document.getElementById(DOM_IDS[`${type.toUpperCase()}_INPUT`]);
    const parts = emailParts[type];

    container.querySelectorAll('.pill-tag').forEach(pill => pill.remove());

    parts.forEach((part, index) => {
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

// --- General UI Getters ---

export function getPowerAutomateUrl() {
    return document.getElementById(DOM_IDS.POWER_AUTOMATE_URL).value.trim();
}

export function getSelectedSheetName() {
    const recipientList = document.getElementById(DOM_IDS.RECIPIENT_LIST);
    const customSheetInput = document.getElementById(DOM_IDS.CUSTOM_SHEET_NAME);
    if (recipientList.value === 'custom') {
        return customSheetInput.value.trim();
    }
    return recipientList.value === 'lda' ? getTodaysLdaSheetName() : 'Master List';
}

export function getEmailTemplateFromDOM() {
    const { emailParts } = getState();
    const quill = getQuill();

    // Ensure special blots have their data up-to-date
    const tempEditorDiv = document.createElement('div');
    if (quill && quill.root) {
        tempEditorDiv.innerHTML = quill.root.innerHTML;
        tempEditorDiv.querySelectorAll('.randomize-tag-wrapper').forEach(tagNode => RandomizeBlot.updateOptions(tagNode));
        tempEditorDiv.querySelectorAll('.condition-tag-wrapper').forEach(tagNode => ConditionBlot.updateOptions(tagNode));
    }


    return {
        from: reconstructPillboxString(emailParts.from),
        subject: reconstructPillboxString(emailParts.subject),
        cc: reconstructPillboxString(emailParts.cc, ';'),
        body: tempEditorDiv.innerHTML
    };
}

function reconstructPillboxString(parts, separator = '') {
    return parts.join(separator);
}

// --- UI Updaters and Status ---

export function updateStatus(message, color = 'gray', isSetupStatus = false) {
    const statusEl = document.getElementById(isSetupStatus ? DOM_IDS.SETUP_STATUS : DOM_IDS.STATUS);
    statusEl.textContent = message;
    statusEl.style.color = color;
}


export function toggleCustomSheetInput() {
    const recipientList = document.getElementById(DOM_IDS.RECIPIENT_LIST);
    const customSheetContainer = document.getElementById(DOM_IDS.CUSTOM_SHEET_CONTAINER);
    customSheetContainer.classList.toggle('hidden', recipientList.value !== 'custom');
}

// --- Modal Content Management ---

export function populatePayloadModal(payload) {
    document.getElementById(DOM_IDS.PAYLOAD_CONTENT).textContent = JSON.stringify(payload, null, 2);
}

export function populateExampleModal(example) {
    document.getElementById(DOM_IDS.EXAMPLE_FROM).textContent = example.from || '[Not Specified]';
    document.getElementById(DOM_IDS.EXAMPLE_TO).textContent = example.to || '[No Email Found]';
    document.getElementById(DOM_IDS.EXAMPLE_CC).textContent = example.cc || '[Not Specified]';
    document.getElementById(DOM_IDS.EXAMPLE_SUBJECT).textContent = example.subject;
    document.getElementById(DOM_IDS.EXAMPLE_BODY).innerHTML = example.body;
}

export function populateSendConfirmationModal(count) {
    document.getElementById(DOM_IDS.SEND_CONFIRM_MESSAGE).textContent = `You are about to send emails to ${count} student(s). Do you want to proceed?`;
}

// --- Templates UI ---

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
            </div>
        `;
        container.appendChild(div);
    });
    
    container.querySelectorAll('.load-template-btn').forEach(btn => {
        btn.onclick = () => handlers.onLoad(btn.dataset.id);
    });
    container.querySelectorAll('.delete-template-btn').forEach(btn => {
        btn.onclick = () => handlers.onDelete(btn.dataset.id);
    });
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

export function loadTemplateIntoForm(template) {
    const { emailParts } = getState();
    const quill = getQuill();

    const parsePills = (str) => str ? str.split(/({[^}]+})/g).filter(part => part) : [];

    emailParts.from = parsePills(template.from);
    emailParts.subject = parsePills(template.subject);
    emailParts.cc = parsePills(template.cc);
    
    renderPills('from');
    renderPills('subject');
    renderPills('cc');

    quill.root.innerHTML = template.body;
}

// --- Custom Parameter UI ---

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

export function addMappingRow(ifValue = '', thenValue = '') {
    const container = document.getElementById(DOM_IDS.PARAM_MAPPING_CONTAINER);
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

export function getCustomParamForm() {
    const mappings = [];
    document.querySelectorAll(`#${DOM_IDS.PARAM_MAPPING_CONTAINER} .mapping-row`).forEach(row => {
        const ifValue = row.querySelector('.mapping-if').value.trim();
        const thenValue = row.querySelector('.mapping-then').value.trim();
        if (ifValue) mappings.push({ if: ifValue, then: thenValue });
    });

    return {
        id: document.getElementById(DOM_IDS.PARAM_EDIT_ID).value,
        name: document.getElementById(DOM_IDS.PARAM_NAME).value.trim(),
        sourceColumn: document.getElementById(DOM_IDS.PARAM_SOURCE_COLUMN).value.trim(),
        defaultValue: document.getElementById(DOM_IDS.PARAM_DEFAULT_VALUE).value.trim(),
        mappings
    };
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
            <div class="mt-2 text-sm">${mappingsHtml}</div>
        `;
        listContainer.appendChild(div);
    });

    listContainer.querySelectorAll('.duplicate-param-btn').forEach(btn => btn.onclick = () => handlers.onDuplicate(btn.dataset.id));
    listContainer.querySelectorAll('.edit-param-btn').forEach(btn => btn.onclick = () => handlers.onEdit(btn.dataset.id));
    listContainer.querySelectorAll('.delete-param-btn').forEach(btn => btn.onclick = () => handlers.onDelete(btn.dataset.id));
}

