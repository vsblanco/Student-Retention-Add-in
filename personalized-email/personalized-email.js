import { findColumnIndex, getTodaysLdaSheetName, getNameParts } from './utils.js';

let powerAutomateConnection = null;
let studentDataCache = [];
let lastFocusedInput = null;
let quill; // To hold the editor instance
let fromParts = [];
let subjectParts = [];
let ccRecipients = [];
let customParameters = [];

// Define Blot classes in a scope accessible by all functions
let RandomizeBlot;
let ConditionBlot;


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
        document.getElementById('create-custom-param-button').onclick = () => showCustomParamModal();

        // Dropdown listener
        document.getElementById('recipient-list').onchange = toggleCustomSheetInput;

        // Modal Close Buttons
        document.getElementById('close-example-modal-button').onclick = () => document.getElementById('example-modal').classList.add('hidden');
        document.getElementById('close-payload-modal-button').onclick = () => document.getElementById('payload-modal').classList.add('hidden');
        document.getElementById('close-templates-modal-button').onclick = () => document.getElementById('templates-modal').classList.add('hidden');
        document.getElementById('cancel-save-template-button').onclick = () => document.getElementById('save-template-modal').classList.add('hidden');
        document.getElementById('cancel-send-button').onclick = () => document.getElementById('send-confirm-modal').classList.add('hidden');
        document.getElementById('cancel-custom-param-button').onclick = () => document.getElementById('custom-param-modal').classList.add('hidden');
        document.getElementById('close-manage-params-button').onclick = () => document.getElementById('manage-custom-params-modal').classList.add('hidden');
        
        // Modal Action Buttons
        document.getElementById('toggle-payload-schema-button').onclick = togglePayloadView;
        document.getElementById('save-current-template-button').onclick = showSaveTemplateModal;
        document.getElementById('confirm-save-template-button').onclick = saveTemplate;
        document.getElementById('confirm-send-button').onclick = executeSend;
        document.getElementById('save-custom-param-button').onclick = saveCustomParameter;
        document.getElementById('add-mapping-button').onclick = () => addMappingRow();
        document.getElementById('manage-custom-params-button').onclick = showManageCustomParamsModal;

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
        
        let Embed = Quill.import('blots/embed');
        let Inline = Quill.import('blots/inline');

        class ParameterBlot extends Inline {
            static create(value) {
                let node = super.create();
                node.setAttribute('contenteditable', 'false');
                node.setAttribute('data-param', value.name);
                node.classList.add('parameter-tag');
                if (value.isCustom) node.classList.add('custom');
                
                // Use innerHTML to structure it
                node.innerHTML = `<span>{${value.name}}</span><span class="parameter-remove-btn" contenteditable="false">×</span>`;

                node.querySelector('.parameter-remove-btn').addEventListener('click', (e) => {
                    e.stopPropagation();
                    const blot = Quill.find(node);
                    if (blot) blot.remove();
                });

                return node;
            }
            static formats(node) {
                return { name: node.getAttribute('data-param'), isCustom: node.classList.contains('custom') };
            }
        }
        ParameterBlot.blotName = 'parameter';
        ParameterBlot.tagName = 'SPAN';
        Quill.register(ParameterBlot);

        // Assign class definitions to the module-scoped variables
        RandomizeBlot = class extends Embed {
            static create(value) {
                const wrapper = document.createElement('span');
                wrapper.classList.add('randomize-tag-wrapper');
                wrapper.setAttribute('contenteditable', 'false');

                const node = document.createElement('span');
                node.classList.add('randomize-tag');
                
                node.innerHTML = `
                    <span>{Randomize}</span>
                    <span class="parameter-remove-btn" contenteditable="false">×</span>
                    <span class="randomize-arrow">&#9660;</span>
                `;
                
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
                
                node.querySelector('.parameter-remove-btn').onclick = (e) => {
                    e.stopPropagation();
                    const blot = Quill.find(wrapper);
                    if (blot) blot.remove();
                };

                node.onclick = (e) => {
                    if (e.target.classList.contains('parameter-remove-btn')) return;
                    e.stopPropagation();
                    const arrow = node.querySelector('.randomize-arrow');
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
        }
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
                node.innerHTML = `
                    <span>{Condition}</span>
                    <span class="parameter-remove-btn" contenteditable="false">×</span>
                    <span class="condition-arrow">&#9660;</span>
                `;

                const panel = document.createElement('div');
                panel.classList.add('condition-panel');
                panel.style.display = 'none';

                const ifClause = document.createElement('div');
                ifClause.classList.add('condition-clause');
                const ifParamInput = document.createElement('input');
                ifParamInput.className = 'condition-input param-input';
                ifParamInput.placeholder = '{Parameter}';
                ifParamInput.value = value.if_param ? `{${value.if_param}}` : '';
                ifParamInput.onfocus = () => { lastFocusedInput = ifParamInput };
                
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
                
                node.querySelector('.parameter-remove-btn').onclick = (e) => {
                    e.stopPropagation();
                    const blot = Quill.find(wrapper);
                    if (blot) blot.remove();
                };

                node.onclick = (e) => {
                    if (e.target.classList.contains('parameter-remove-btn')) return;
                    e.stopPropagation();
                    const arrow = node.querySelector('.condition-arrow');
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
        }
        ConditionBlot.blotName = 'condition';
        ConditionBlot.tagName = 'SPAN';
        Quill.register(ConditionBlot);
        
        setupPillboxInput('from');
        setupPillboxInput('subject');
        setupPillboxInput('cc');
        
        quill.on('selection-change', (range) => {
            if (range) {
                lastFocusedInput = quill;
            }
        });

        document.addEventListener('click', (e) => {
             if (!e.target.closest('.randomize-tag-wrapper') && !e.target.closest('.condition-tag-wrapper')) {
                document.querySelectorAll('.randomize-panel, .condition-panel').forEach(p => p.style.display = 'none');
                document.querySelectorAll('.randomize-arrow, .condition-arrow').forEach(a => a.classList.remove('open'));
            }
        });


        loadCustomParameters().then(populateParameterButtons);
        checkConnection();
        toggleCustomSheetInput();
    }
});

async function populateParameterButtons() {
    const standardContainer = document.getElementById('standard-parameter-buttons');
    const customContainer = document.getElementById('custom-parameter-buttons');
    const customSection = document.getElementById('custom-parameters-section');

    standardContainer.innerHTML = '';
    customContainer.innerHTML = '';

    const createButton = (param, container, isCustom = false) => {
        const button = document.createElement('button');
        if (isCustom) {
            button.className = 'px-2 py-1 bg-blue-100 text-blue-800 text-xs rounded hover:bg-blue-200';
        } else {
            button.className = 'px-2 py-1 bg-gray-200 text-gray-800 text-xs rounded hover:bg-gray-300';
        }
        button.textContent = `{${param}}`;
        button.onmousedown = (e) => {
            e.preventDefault(); 
            insertParameter(`{${param}}`);
        };
        container.appendChild(button);
    };

    standardParameters.forEach(param => createButton(param, standardContainer, false));

    if (customParameters.length > 0) {
        customSection.classList.remove('hidden');
        customParameters.forEach(param => createButton(param.name, customContainer, true));
    } else {
        customSection.classList.add('hidden');
    }

    document.getElementById('randomize-parameter-button').onmousedown = (e) => { e.preventDefault(); insertParameter('{Randomize}'); };
    document.getElementById('condition-parameter-button').onmousedown = (e) => { e.preventDefault(); insertParameter('{Condition}'); };
}

function insertParameter(param) {
    const paramName = param.replace(/[{}]/g, '');

    // Check if the last focused element is the special input inside the condition panel
    if (lastFocusedInput && lastFocusedInput.classList && lastFocusedInput.classList.contains('param-input')) {
        lastFocusedInput.value = param;
        // Manually trigger the input event to save the change to the blot's dataset
        lastFocusedInput.dispatchEvent(new Event('input', { bubbles: true }));
    } else if (lastFocusedInput && ['email-from-input', 'email-subject-input', 'email-cc-input'].includes(lastFocusedInput.id)) {
        // Check if the last focused element is one of our main pillbox inputs
        const type = lastFocusedInput.id.replace('email-', '').replace('-input', ''); // from, subject, or cc
        addPill(type, param);
    } else {
        // Default to Quill editor if it's not a known input or if Quill was last focused
        const range = quill.getSelection(true); // Get current position, or end
        if (paramName === 'Randomize') {
            quill.insertEmbed(range.index, 'randomize', { options: [''] }, Quill.sources.USER);
        } else if (paramName === 'Condition') {
            quill.insertEmbed(range.index, 'condition', {}, Quill.sources.USER);
        } else {
            const isCustom = customParameters.some(p => p.name === paramName);
            quill.insertEmbed(range.index, 'parameter', { name: paramName, isCustom }, Quill.sources.USER);
        }
        quill.setSelection(range.index + 1, Quill.sources.USER);
    }
}


function toggleCustomSheetInput() {
    const recipientList = document.getElementById('recipient-list');
    const customSheetContainer = document.getElementById('custom-sheet-container');
    customSheetContainer.classList.toggle('hidden', recipientList.value !== 'custom');
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
                };

                customParameters.forEach(param => {
                    const colIndex = customParamIndices[param.name];
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
    const tempDiv = document.createElement('div');
    tempDiv.innerHTML = bodyHtml;
    
    tempDiv.querySelectorAll('.parameter-tag').forEach(tag => {
        const paramName = tag.getAttribute('data-param');
        const value = data[paramName] ?? tag.innerText;
        tag.replaceWith(document.createTextNode(value));
    });

    tempDiv.querySelectorAll('.randomize-tag-wrapper').forEach(tagWrapper => {
        const options = JSON.parse(tagWrapper.dataset.options || '[]');
        const choice = options.length > 0 ? options[Math.floor(Math.random() * options.length)] : '';
        tagWrapper.replaceWith(document.createTextNode(choice));
    });

    tempDiv.querySelectorAll('.condition-tag-wrapper').forEach(tagWrapper => {
        const condition = JSON.parse(tagWrapper.dataset.condition || '{}');
        const studentValue = data[condition.if_param];
        let showText = false;
        
        const val1 = isNaN(parseFloat(studentValue)) ? String(studentValue).toLowerCase() : parseFloat(studentValue);
        const val2 = isNaN(parseFloat(condition.if_value)) ? String(condition.if_value).toLowerCase() : parseFloat(condition.if_value);

        switch(condition.operator) {
            case '=': showText = val1 == val2; break;
            case '>': showText = val1 > val2; break;
            case '>=': showText = val1 >= val2; break;
            case '<': showText = val1 < val2; break;
            case '<=': showText = val1 <= val2; break;
        }
        
        const outputText = showText ? renderTemplate(condition.then_text, data) : '';
        tagWrapper.replaceWith(document.createTextNode(outputText));
    });


    return tempDiv.innerHTML;
};

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
        
        const fromTemplate = reconstructPillboxString(fromParts);
        const subjectTemplate = reconstructPillboxString(subjectParts);
        const ccTemplate = reconstructPillboxString(ccRecipients, ';');
        const bodyTemplate = quill.root.innerHTML;

        document.getElementById('example-from').textContent = renderTemplate(fromTemplate, randomStudent) || '[Not Specified]';
        document.getElementById('example-to').textContent = randomStudent.StudentEmail || '[No Email Found]';
        document.getElementById('example-cc').textContent = renderTemplate(ccTemplate, randomStudent) || '[Not Specified]';
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

        const fromTemplate = reconstructPillboxString(fromParts);
        const subjectTemplate = reconstructPillboxString(subjectParts);
        const ccTemplate = reconstructPillboxString(ccRecipients, ';');
        
        const tempEditorDiv = document.createElement('div');
        tempEditorDiv.innerHTML = quill.root.innerHTML;
        tempEditorDiv.querySelectorAll('.randomize-tag-wrapper, .condition-tag-wrapper').forEach(tagNode => {
            if (tagNode.classList.contains('randomize-tag-wrapper')) RandomizeBlot.updateOptions(tagNode);
            if (tagNode.classList.contains('condition-tag-wrapper')) ConditionBlot.updateOptions(tagNode);
        });
        const bodyTemplate = tempEditorDiv.innerHTML;


        const payload = studentDataCache.map(student => ({
            from: renderTemplate(fromTemplate, student),
            to: student.StudentEmail || '',
            cc: renderTemplate(ccTemplate, student),
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

    const fromTemplate = reconstructPillboxString(fromParts);
    const subjectTemplate = reconstructPillboxString(subjectParts);
    const ccTemplate = reconstructPillboxString(ccRecipients, ';');
    
    const tempEditorDiv = document.createElement('div');
    tempEditorDiv.innerHTML = quill.root.innerHTML;
    tempEditorDiv.querySelectorAll('.randomize-tag-wrapper, .condition-tag-wrapper').forEach(tagNode => {
         if (tagNode.classList.contains('randomize-tag-wrapper')) RandomizeBlot.updateOptions(tagNode);
         if (tagNode.classList.contains('condition-tag-wrapper')) ConditionBlot.updateOptions(tagNode);
    });
    const bodyTemplate = tempEditorDiv.innerHTML;


    const payload = studentDataCache.map(student => ({
        from: renderTemplate(fromTemplate, student),
        to: student.StudentEmail || '',
        cc: renderTemplate(ccTemplate, student),
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

    const tempEditorDiv = document.createElement('div');
    tempEditorDiv.innerHTML = quill.root.innerHTML;
    tempEditorDiv.querySelectorAll('.randomize-tag-wrapper, .condition-tag-wrapper').forEach(tagNode => {
         if (tagNode.classList.contains('randomize-tag-wrapper')) RandomizeBlot.updateOptions(tagNode);
         if (tagNode.classList.contains('condition-tag-wrapper')) ConditionBlot.updateOptions(tagNode);
         const panel = tagNode.querySelector('.randomize-panel, .condition-panel');
         if(panel) panel.remove();
    });

    const newTemplate = {
        id: 'template_' + new Date().getTime(),
        name: name,
        author: author,
        timestamp: new Date().toISOString(),
        from: reconstructPillboxString(fromParts),
        subject: reconstructPillboxString(subjectParts),
        cc: reconstructPillboxString(ccRecipients),
        body: tempEditorDiv.innerHTML
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
        fromParts = parseStringToPills(template.from || '');
        subjectParts = parseStringToPills(template.subject || '');
        ccRecipients = parseStringToPills(template.cc || '');
        
        renderPills('from');
        renderPills('subject');
        renderPills('cc');

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

function showCustomParamModal(paramToEdit = null) {
    const modalTitle = document.getElementById('custom-param-modal-title');
    const nameInput = document.getElementById('param-name');
    const sourceColumnInput = document.getElementById('param-source-column');
    const defaultValueInput = document.getElementById('param-default-value');
    const mappingContainer = document.getElementById('param-mapping-container');
    const editIdInput = document.getElementById('param-edit-id');

    mappingContainer.innerHTML = '';
    document.getElementById('save-param-status').textContent = '';

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
    
    document.getElementById('custom-param-modal').classList.remove('hidden');
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

async function saveCustomParameter() {
    const status = document.getElementById('save-param-status');
    const nameInput = document.getElementById('param-name');
    const name = nameInput.value.trim();
    const paramId = document.getElementById('param-edit-id').value;

    if (!/^[a-zA-Z0-9]+$/.test(name)) {
        status.textContent = 'Name must be alphanumeric with no spaces.';
        status.style.color = 'red';
        return;
    }
    const existingParam = customParameters.find(p => p.name.toLowerCase() === name.toLowerCase());
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

    const paramData = { name, sourceColumn, defaultValue, mappings };

    status.textContent = 'Saving...';
    status.style.color = 'gray';

    let currentParams = await getCustomParameters();
    if (paramId) {
        const index = currentParams.findIndex(p => p.id === paramId);
        if (index > -1) currentParams[index] = { ...currentParams[index], ...paramData };
    } else {
        paramData.id = 'sparam_' + new Date().getTime();
        currentParams.push(paramData);
    }
    
    await saveCustomParameters(currentParams);
    await loadCustomParameters();
    await populateParameterButtons();

    status.textContent = 'Parameter saved successfully!';
    status.style.color = 'green';
    setTimeout(() => {
        document.getElementById('custom-param-modal').classList.add('hidden');
    }, 1500);
}

async function showManageCustomParamsModal() {
    document.getElementById('custom-param-modal').classList.add('hidden');
    const listContainer = document.getElementById('manage-params-list');
    listContainer.innerHTML = '<p class="text-gray-500">Loading...</p>';
    document.getElementById('manage-custom-params-modal').classList.remove('hidden');

    const params = await getCustomParameters();
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

    listContainer.querySelectorAll('.duplicate-param-btn').forEach(btn => btn.onclick = () => duplicateCustomParameter(btn.dataset.id));
    listContainer.querySelectorAll('.edit-param-btn').forEach(btn => btn.onclick = () => editCustomParameter(btn.dataset.id));
    listContainer.querySelectorAll('.delete-param-btn').forEach(btn => btn.onclick = () => deleteCustomParameter(btn.dataset.id));
}

function editCustomParameter(paramId) {
    const param = customParameters.find(p => p.id === paramId);
    if (param) {
        document.getElementById('manage-custom-params-modal').classList.add('hidden');
        showCustomParamModal(param);
    }
}

async function duplicateCustomParameter(paramId) {
    let params = await getCustomParameters();
    const paramToDuplicate = params.find(p => p.id === paramId);
    if (!paramToDuplicate) return;

    const newParam = JSON.parse(JSON.stringify(paramToDuplicate));
    let newName = `${newParam.name}Copy`;
    const allParamNames = [...standardParameters, ...params.map(p => p.name)];
    while (allParamNames.includes(newName)) newName = `${newName}Copy`;
    newParam.name = newName;
    newParam.id = 'sparam_' + new Date().getTime();
    params.push(newParam);
    await saveCustomParameters(params);
    await loadCustomParameters(); 
    await populateParameterButtons();
    await showManageCustomParamsModal();
}

async function deleteCustomParameter(paramId) {
    let params = await getCustomParameters();
    params = params.filter(p => p.id !== paramId);
    await saveCustomParameters(params);
    await loadCustomParameters();
    await populateParameterButtons();
    await showManageCustomParamsModal();
}

// --- Pillbox Input Functions ---
function setupPillboxInput(type) {
    const container = document.getElementById(`email-${type}-container`);
    const input = document.getElementById(`email-${type}-input`);

    container.addEventListener('click', () => {
        lastFocusedInput = input;
        input.focus();
    });

    input.addEventListener('focus', () => lastFocusedInput = input);

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
}

function addPill(type, text) {
    if (!text) return;
    const parts = getPartsArray(type);
    parts.push(text);
    renderPills(type);
}

function removePill(type, index) {
    const parts = getPartsArray(type);
    parts.splice(index, 1);
    renderPills(type);
}

function removeLastPill(type) {
    const parts = getPartsArray(type);
    if (parts.length > 0) {
        parts.pop();
        renderPills(type);
    }
}

function getPartsArray(type) {
    if (type === 'from') return fromParts;
    if (type === 'subject') return subjectParts;
    if (type === 'cc') return ccRecipients;
    return [];
}

function renderPills(type) {
    const container = document.getElementById(`email-${type}-container`);
    const input = document.getElementById(`email-${type}-input`);
    const parts = getPartsArray(type);
    
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

function reconstructPillboxString(parts, separator = '') {
    return parts.join(separator);
}

function parseStringToPills(str) {
    if (!str) return [];
    // This regex splits by parameters, keeping the parameters in the array.
    const regex = /({[^}]+})/g;
    return str.split(regex).filter(part => part); // filter out empty strings
}
