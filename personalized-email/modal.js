// V-1.2 - 2025-09-30 - 5:40 PM EDT
import { PAYLOAD_SCHEMA, MAPPING_OPERATORS, MINI_QUILL_EDITOR_CONFIG } from './constants.js';

export default class ModalManager {
    constructor(appContext) {
        this.appContext = appContext;
        this.currentExampleIndex = 0;
        this.isEditingTemplate = false;
        this.templateToEdit = null;
        this.isEditingParam = false;
        this.paramToEdit = null;
        this.paramMappingQuillInstances = [];

        // Bind methods
        this.showSendConfirmModal = this.showSendConfirmModal.bind(this);
        this.showRecipientModal = this.showRecipientModal.bind(this);
        
        // Setup UI Listeners
        this._setupEventListeners();
    }

    _setupEventListeners() {
        // Example Modal
        document.getElementById('show-example-button').onclick = () => this.showExampleModal();
        document.getElementById('close-example-modal-button').onclick = () => this._hideModal('example-modal');
        document.getElementById('prev-student-button').onclick = () => this._navigateExample(-1);
        document.getElementById('next-student-button').onclick = () => this._navigateExample(1);
        document.getElementById('random-student-button').onclick = () => this._showRandomExample();
        document.getElementById('search-student-button').onclick = () => this._toggleSearch();
        document.getElementById('example-search-input').addEventListener('input', (e) => this._filterStudents(e.target.value));

        // Payload Modal
        document.getElementById('close-payload-modal-button').onclick = () => this._hideModal('payload-modal');
        document.getElementById('toggle-payload-schema-button').onclick = (e) => this._togglePayloadView(e.target);

        // Templates Modals
        document.getElementById('templates-button').onclick = () => this.showTemplatesModal();
        document.getElementById('close-templates-modal-button').onclick = () => this._hideModal('templates-modal');
        document.getElementById('save-current-template-button').onclick = () => this.showSaveTemplateModal();
        document.getElementById('cancel-save-template-button').onclick = () => this._hideModal('save-template-modal');
        document.getElementById('confirm-save-template-button').onclick = () => this._saveTemplate();
        document.getElementById('delete-template-button').onclick = () => this._deleteTemplate();
        
        // Custom Parameter Modals
        document.getElementById('create-custom-param-button').onclick = () => this.showCustomParamModal();
        document.getElementById('cancel-custom-param-button').onclick = () => this._hideModal('custom-param-modal');
        document.getElementById('save-custom-param-button').onclick = () => this._saveCustomParameter();
        document.getElementById('manage-custom-params-button').onclick = () => this.showManageParamsModal();
        document.getElementById('close-manage-custom-params-button').onclick = () => this._hideModal('manage-custom-params-modal');
        document.getElementById('add-mapping-button').onclick = (e) => { e.preventDefault(); this._addMappingRow(); };
        document.getElementById('logic-type-dropdown').onchange = (e) => this._handleLogicTypeChange(e.target.value);
        
        // Send Confirm Modal
        document.getElementById('cancel-send-button').onclick = () => this._hideModal('send-confirm-modal');
        document.getElementById('confirm-send-button').onclick = this.appContext.executeSend;
        
        // Recipient Modal
        document.getElementById('cancel-recipient-modal-button').onclick = () => this._hideModal('recipient-modal');
        document.getElementById('confirm-recipient-modal-button').onclick = () => this._confirmRecipientSelection();
        
        const radioButtons = document.querySelectorAll('input[name="recipient-source"]');
        radioButtons.forEach(radio => {
            radio.addEventListener('change', () => this._handleRecipientSourceChange());
        });
    }

    _showModal(id) {
        document.getElementById(id).classList.remove('hidden');
    }

    _hideModal(id) {
        document.getElementById(id).classList.add('hidden');
    }
    
    // --- Recipient Modal Logic ---
    showRecipientModal() {
        const { type, customSheetName, excludeDNC, excludeFillColor } = this.appContext.recipientSelection;
        document.querySelector(`input[name="recipient-source"][value="${type}"]`).checked = true;
        document.getElementById('recipient-custom-sheet-name').value = customSheetName;
        document.getElementById('exclude-dnc-toggle').checked = excludeDNC;
        document.getElementById('exclude-fill-color-toggle').checked = excludeFillColor;
        
        this._handleRecipientSourceChange(); // Update UI based on initial state
        this._showModal('recipient-modal');
    }

    _handleRecipientSourceChange() {
        const customSheetContainer = document.getElementById('recipient-custom-sheet-container');
        const confirmButton = document.getElementById('confirm-recipient-modal-button');
        const customRadio = document.getElementById('recipient-custom');
        const customSheetInput = document.getElementById('recipient-custom-sheet-name');

        if (customRadio.checked) {
            customSheetContainer.classList.remove('hidden');
            confirmButton.disabled = customSheetInput.value.trim() === '';
            customSheetInput.oninput = () => {
                confirmButton.disabled = customSheetInput.value.trim() === '';
            };
        } else {
            customSheetContainer.classList.add('hidden');
            confirmButton.disabled = false;
        }
    }

    async _confirmRecipientSelection() {
        const status = document.getElementById('recipient-modal-status');
        const confirmButton = document.getElementById('confirm-recipient-modal-button');
        confirmButton.disabled = true;
        status.textContent = 'Fetching and counting students...';
        status.style.color = 'gray';

        const selectedType = document.querySelector('input[name="recipient-source"]:checked').value;
        const customSheetName = document.getElementById('recipient-custom-sheet-name').value;
        const excludeDNC = document.getElementById('exclude-dnc-toggle').checked;
        const excludeFillColor = document.getElementById('exclude-fill-color-toggle').checked;

        const newSelection = {
            type: selectedType,
            customSheetName: customSheetName,
            excludeDNC: excludeDNC,
            excludeFillColor: excludeFillColor
        };

        try {
            const students = await this.appContext.getStudentDataCore(newSelection);
            this.appContext.updateRecipientSelection(newSelection, students.length);
            status.textContent = `Found ${students.length} students.`;
            status.style.color = 'green';
            setTimeout(() => {
                this._hideModal('recipient-modal');
                status.textContent = '';
                confirmButton.disabled = false;
            }, 1500);
        } catch (error) {
            status.textContent = error.userFacingMessage || 'An error occurred.';
            status.style.color = 'red';
            confirmButton.disabled = false;
        }
    }


    // --- Send Confirm Modal ---
    showSendConfirmModal() {
        if (this.appContext.studentDataCache.length > 0) {
            document.getElementById('send-confirm-message').textContent = `You are about to send personalized emails to ${this.appContext.studentDataCache.length} students. Are you sure?`;
            this._showModal('send-confirm-modal');
        } else {
            const status = document.getElementById('status');
            status.textContent = 'Please select recipients before sending.';
            status.style.color = 'orange';
            setTimeout(() => status.textContent = '', 3000);
        }
    }

    // --- Example/Preview Modal Logic ---
    async showExampleModal() {
        if (this.appContext.studentDataCache.length === 0) {
            try {
                await this.appContext.getStudentDataWithUI();
                if (this.appContext.studentDataCache.length > 0) {
                    this._renderExample();
                    this._showModal('example-modal');
                }
            } catch (error) {
                // Error message is already displayed by getStudentDataWithUI
                return;
            }
        } else {
            this._renderExample();
            this._showModal('example-modal');
        }
    }
    
    _renderExample() {
        const student = this.appContext.studentDataCache[this.currentExampleIndex];
        if (!student) {
            document.getElementById('example-student-counter').textContent = "No student found";
            return;
        }

        const { renderTemplate, renderCCTemplate, quill, ccRecipients } = this.appContext;
        const fromTemplate = document.getElementById('email-from').value;
        const subjectTemplate = document.getElementById('email-subject').value;
        const bodyTemplate = quill.root.innerHTML;

        document.getElementById('example-from').textContent = renderTemplate(fromTemplate, student);
        document.getElementById('example-to').textContent = student.StudentEmail || '[No Email]';
        document.getElementById('example-cc').textContent = renderCCTemplate(ccRecipients, student) || '[None]';
        document.getElementById('example-subject').textContent = renderTemplate(subjectTemplate, student);
        document.getElementById('example-body').innerHTML = renderTemplate(bodyTemplate, student);
        
        this._updateExampleCounter();
    }

    _updateExampleCounter() {
        const total = this.appContext.studentDataCache.length;
        document.getElementById('example-student-counter').textContent = `Student ${this.currentExampleIndex + 1} of ${total}`;
        document.getElementById('prev-student-button').disabled = this.currentExampleIndex === 0;
        document.getElementById('next-student-button').disabled = this.currentExampleIndex >= total - 1;
    }
    
    _navigateExample(direction) {
        const newIndex = this.currentExampleIndex + direction;
        if (newIndex >= 0 && newIndex < this.appContext.studentDataCache.length) {
            this.currentExampleIndex = newIndex;
            this._renderExample();
        }
    }

    _showRandomExample() {
        const total = this.appContext.studentDataCache.length;
        if (total > 0) {
            this.currentExampleIndex = Math.floor(Math.random() * total);
            this._renderExample();
        }
    }

    _toggleSearch() {
        document.getElementById('example-search-container').classList.toggle('hidden');
    }

    _filterStudents(query) {
        const resultsContainer = document.getElementById('example-search-results');
        if (!query) {
            resultsContainer.classList.add('hidden');
            return;
        }

        const filtered = this.appContext.studentDataCache.filter(student =>
            student.StudentName.toLowerCase().includes(query.toLowerCase())
        ).slice(0, 5); // Limit results for performance

        resultsContainer.innerHTML = '';
        if (filtered.length > 0) {
            filtered.forEach(student => {
                const item = document.createElement('div');
                item.className = 'px-3 py-2 cursor-pointer hover:bg-gray-100';
                item.textContent = student.StudentName;
                item.onclick = () => {
                    this.currentExampleIndex = this.appContext.studentDataCache.indexOf(student);
                    this._renderExample();
                    resultsContainer.classList.add('hidden');
                    document.getElementById('example-search-input').value = '';
                };
                resultsContainer.appendChild(item);
            });
            resultsContainer.classList.remove('hidden');
        } else {
            resultsContainer.classList.add('hidden');
        }
    }


    // --- Payload Modal ---
    showPayloadModal() {
        const { renderTemplate, renderCCTemplate, quill, ccRecipients, studentDataCache } = this.appContext;
        
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

        document.getElementById('payload-content').textContent = JSON.stringify(payload, null, 2);
        document.getElementById('schema-content').textContent = JSON.stringify(PAYLOAD_SCHEMA, null, 2);
        this._showModal('payload-modal');
    }

    _togglePayloadView(button) {
        const payloadContent = document.getElementById('payload-content');
        const schemaContent = document.getElementById('schema-content');
        const title = document.getElementById('payload-modal-title');

        const isSchemaVisible = !schemaContent.classList.contains('hidden');
        if (isSchemaVisible) {
            schemaContent.classList.add('hidden');
            payloadContent.classList.remove('hidden');
            title.textContent = 'Request Payload';
            button.textContent = 'Show Schema';
        } else {
            schemaContent.classList.remove('hidden');
            payloadContent.classList.add('hidden');
            title.textContent = 'Payload JSON Schema';
            button.textContent = 'Show Payload';
        }
    }


    // --- Template Management Modals ---
    async showTemplatesModal() {
        const templates = await this.appContext.getTemplates();
        const container = document.getElementById('templates-list-container');
        container.innerHTML = '';

        if (templates.length === 0) {
            container.innerHTML = '<p class="text-sm text-gray-500 text-center">No saved templates found.</p>';
        } else {
            templates.forEach(template => {
                const div = document.createElement('div');
                div.className = 'p-3 border rounded-md hover:bg-gray-50';
                div.innerHTML = `
                    <div class="flex justify-between items-start">
                        <div>
                            <p class="font-semibold text-gray-800">${template.name}</p>
                            <p class="text-xs text-gray-500">By: ${template.author || 'Unknown'}</p>
                        </div>
                        <div class="flex gap-2">
                            <button data-action="load" class="text-sm text-blue-600 hover:underline">Load</button>
                            <button data-action="edit" class="text-sm text-gray-600 hover:underline">Edit</button>
                        </div>
                    </div>
                `;
                div.querySelector('button[data-action="load"]').onclick = () => this._loadTemplate(template);
                div.querySelector('button[data-action="edit"]').onclick = () => this.showSaveTemplateModal(template);
                container.appendChild(div);
            });
        }
        this._showModal('templates-modal');
    }

    showSaveTemplateModal(template = null) {
        this.isEditingTemplate = !!template;
        this.templateToEdit = template;

        document.getElementById('save-template-modal-title').textContent = this.isEditingTemplate ? 'Edit Template' : 'Save as Template';
        document.getElementById('template-name').value = this.isEditingTemplate ? template.name : '';
        document.getElementById('template-author').value = this.isEditingTemplate ? template.author : '';
        document.getElementById('delete-template-button').classList.toggle('hidden', !this.isEditingTemplate);
        
        this._hideModal('templates-modal');
        this._showModal('save-template-modal');
    }

    async _saveTemplate() {
        const name = document.getElementById('template-name').value.trim();
        const author = document.getElementById('template-author').value.trim();
        if (!name) {
            document.getElementById('save-template-status').textContent = 'Template name is required.';
            return;
        }

        const templates = await this.appContext.getTemplates();
        const newTemplate = {
            id: this.isEditingTemplate ? this.templateToEdit.id : `template-${Date.now()}`,
            name,
            author,
            from: document.getElementById('email-from').value,
            cc: this.appContext.ccRecipients,
            subject: document.getElementById('email-subject').value,
            body: this.appContext.quill.root.innerHTML
        };

        if (this.isEditingTemplate) {
            const index = templates.findIndex(t => t.id === this.templateToEdit.id);
            templates[index] = newTemplate;
        } else {
            templates.push(newTemplate);
        }

        await this.appContext.saveTemplates(templates);
        this._hideModal('save-template-modal');
    }

    async _deleteTemplate() {
        if (!this.templateToEdit) return;
        let templates = await this.appContext.getTemplates();
        templates = templates.filter(t => t.id !== this.templateToEdit.id);
        await this.appContext.saveTemplates(templates);
        this._hideModal('save-template-modal');
    }
    
    _loadTemplate(template) {
        document.getElementById('email-from').value = template.from;
        document.getElementById('email-subject').value = template.subject;
        this.appContext.quill.root.innerHTML = template.body;
        
        // Handle CC recipients
        this.appContext.ccRecipients.length = 0;
        if(template.cc && Array.isArray(template.cc)) {
             this.appContext.ccRecipients.push(...template.cc);
        }
        this.appContext.renderCCPills();
        
        this._hideModal('templates-modal');
    }
    
    // --- Custom Parameter Modals ---
    showCustomParamModal(param = null) {
        this.isEditingParam = !!param;
        this.paramToEdit = param;

        document.getElementById('custom-param-modal-title').textContent = this.isEditingParam ? 'Edit Custom Parameter' : 'Create Custom Parameter';
        document.getElementById('param-name').value = this.isEditingParam ? param.name : '';
        document.getElementById('param-source-column').value = this.isEditingParam ? param.sourceColumn : '';

        // Reset and populate logic type
        const logicDropdown = document.getElementById('logic-type-dropdown');
        logicDropdown.value = this.isEditingParam ? (param.logicType || '') : '';
        this._handleLogicTypeChange(logicDropdown.value); // This will show/hide the correct container

        // Populate mappings if they exist
        const mappingContainer = document.getElementById('param-mapping-container');
        mappingContainer.innerHTML = '';
        this.paramMappingQuillInstances = []; // Clear old Quill instances
        if (this.isEditingParam && param.logicType === 'value-mapping' && param.mappings) {
            param.mappings.forEach(m => this._addMappingRow(m));
        }

        this._showModal('custom-param-modal');
    }

    async _saveCustomParameter() {
        const name = document.getElementById('param-name').value.trim();
        const sourceColumn = document.getElementById('param-source-column').value.trim();
        const logicType = document.getElementById('logic-type-dropdown').value;

        if (!name || !sourceColumn) {
            document.getElementById('save-param-status').textContent = 'Parameter Name and Source Column are required.';
            return;
        }
        if (!/^[a-zA-Z0-9_]+$/.test(name)) {
            document.getElementById('save-param-status').textContent = 'Parameter Name can only contain letters, numbers, and underscores.';
            return;
        }

        const newParam = {
            id: this.isEditingParam ? this.paramToEdit.id : `param-${Date.now()}`,
            name,
            sourceColumn,
            logicType: logicType || null
        };

        if (logicType === 'value-mapping') {
            newParam.mappings = [];
            const mappingRows = document.querySelectorAll('.mapping-row');
            for (let i = 0; i < mappingRows.length; i++) {
                const row = mappingRows[i];
                const ifValue = row.querySelector('.mapping-if').value;
                const thenValue = this.paramMappingQuillInstances[i].root.innerHTML;
                if (ifValue) {
                    newParam.mappings.push({
                        if: ifValue,
                        operator: row.querySelector('.mapping-operator').value,
                        then: thenValue
                    });
                }
            }
        }

        let params = await this.appContext.getCustomParameters();
        if (this.isEditingParam) {
            const index = params.findIndex(p => p.id === this.paramToEdit.id);
            params[index] = newParam;
        } else {
            params.push(newParam);
        }

        await this.appContext.saveCustomParameters(params);
        await this.appContext.loadCustomParameters(); // Reload parameters
        this.appContext.populateParameterButtons(); // Refresh UI
        this._hideModal('custom-param-modal');
    }

    async showManageParamsModal() {
        this._hideModal('custom-param-modal');
        const params = await this.appContext.getCustomParameters();
        const container = document.getElementById('manage-custom-params-list');
        container.innerHTML = '';

        if (params.length === 0) {
            container.innerHTML = '<p class="text-sm text-gray-500 text-center">No custom parameters found.</p>';
        } else {
            params.forEach(param => {
                const div = document.createElement('div');
                div.className = 'p-3 border-b flex justify-between items-center';
                div.innerHTML = `
                    <div>
                        <p class="font-semibold text-gray-800">{${param.name}}</p>
                        <p class="text-xs text-gray-500">Source: ${param.sourceColumn}</p>
                    </div>
                    <div>
                        <button data-action="edit" class="text-sm text-blue-600 hover:underline">Edit</button>
                        <button data-action="delete" class="ml-4 text-sm text-red-600 hover:underline">Delete</button>
                    </div>
                `;
                div.querySelector('button[data-action="edit"]').onclick = () => {
                    this._hideModal('manage-custom-params-modal');
                    this.showCustomParamModal(param);
                };
                div.querySelector('button[data-action="delete"]').onclick = async () => {
                    let currentParams = await this.appContext.getCustomParameters();
                    currentParams = currentParams.filter(p => p.id !== param.id);
                    await this.appContext.saveCustomParameters(currentParams);
                    await this.appContext.loadCustomParameters();
                    this.appContext.populateParameterButtons();
                    this.showManageParamsModal(); // Refresh the list
                };
                container.appendChild(div);
            });
        }
        this._showModal('manage-custom-params-modal');
    }

    _handleLogicTypeChange(value) {
        document.getElementById('value-mapping-logic-container').classList.toggle('hidden', value !== 'value-mapping');
        document.getElementById('custom-script-logic-container').classList.toggle('hidden', value !== 'custom-script');
    }

    _addMappingRow(mapping = { if: '', operator: 'eq', then: '' }) {
        const container = document.getElementById('param-mapping-container');
        const div = document.createElement('div');
        div.className = 'mapping-row flex items-center gap-2 p-2 border rounded-md bg-gray-50';

        // Create a unique ID for the Quill editor container
        const quillEditorId = `quill-editor-${this.paramMappingQuillInstances.length}`;

        div.innerHTML = `
            <div class="flex-1">
                <label class="text-xs text-gray-500">If source value...</label>
                <div class="flex items-center gap-1 mt-1">
                    <select class="mapping-operator w-full text-xs border-gray-300 rounded-md shadow-sm">
                        ${MAPPING_OPERATORS.map(op => `<option value="${op.value}" ${mapping.operator === op.value ? 'selected' : ''}>${op.text}</option>`).join('')}
                    </select>
                    <input type="text" class="mapping-if w-full text-xs border-gray-300 rounded-md shadow-sm" value="${mapping.if}">
                </div>
            </div>
            <div class="flex-1">
                <label class="text-xs text-gray-500">...then use this value instead:</label>
                <div id="${quillEditorId}" class="mapping-then mt-1 bg-white rounded-md mini-quill-editor"></div>
            </div>
            <button class="remove-mapping-btn self-end mb-1 p-1 text-red-500 hover:text-red-700">&times;</button>
        `;

        container.appendChild(div);
        
        // Initialize Quill on the new element
        const quillInstance = new Quill(`#${quillEditorId}`, MINI_QUILL_EDITOR_CONFIG);
        quillInstance.root.innerHTML = mapping.then;
        this.paramMappingQuillInstances.push(quillInstance);

        div.querySelector('.remove-mapping-btn').onclick = (e) => {
            e.preventDefault();
            const index = Array.from(container.children).indexOf(div);
            this.paramMappingQuillInstances.splice(index, 1);
            div.remove();
        };
    }
}

