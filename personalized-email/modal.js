// V-2.4 - 2025-10-01 - 1:05 PM EDT
export default class ModalManager {
    constructor(appContext) {
        this.appContext = appContext;
        this.currentExampleIndex = 0;
        this.editingTemplate = null; 
        this.editingParam = null;
        this.setupEventListeners();
        this.currentPayload = null;
        this.currentSchema = {
            "type": "array",
            "items": {
                "type": "object",
                "properties": {
                    "from": { "type": "string", "description": "The sender's email address." },
                    "to": { "type": "string", "description": "The recipient's email address." },
                    "cc": { "type": "string", "description": "CC recipients, separated by semicolons." },
                    "subject": { "type": "string", "description": "The subject line of the email." },
                    "body": { "type": "string", "description": "The HTML body of the email." }
                },
                "required": ["from", "to", "subject", "body"]
            }
        };
    }

    setupEventListeners() {
        // Example Modal
        document.getElementById('show-example-button').onclick = () => this.showExampleModal();
        document.getElementById('close-example-modal-button').onclick = () => this.hideModal('example-modal');
        document.getElementById('prev-student-button').onclick = () => this.showAdjacentStudent(this.currentExampleIndex - 1);
        document.getElementById('next-student-button').onclick = () => this.showAdjacentStudent(this.currentExampleIndex + 1);
        document.getElementById('random-student-button').onclick = () => this.showRandomStudent();
        
        // Example search
        const searchButton = document.getElementById('search-student-button');
        const searchContainer = document.getElementById('example-search-container');
        const searchInput = document.getElementById('example-search-input');
        const searchResults = document.getElementById('example-search-results');

        searchButton.onclick = () => {
            searchContainer.classList.toggle('hidden');
            if (!searchContainer.classList.contains('hidden')) {
                searchInput.focus();
            }
        };

        searchInput.addEventListener('input', () => {
            const query = searchInput.value.toLowerCase();
            if (query.length < 2) {
                searchResults.classList.add('hidden');
                return;
            }
            const filteredStudents = this.appContext.studentDataCache.filter(s => 
                s.StudentName.toLowerCase().includes(query)
            );
            this.renderSearchResults(filteredStudents);
        });
        
        document.addEventListener('click', (e) => {
            if (!searchContainer.contains(e.target) && e.target !== searchButton && !searchButton.contains(e.target)) {
                 searchContainer.classList.add('hidden');
                 searchResults.classList.add('hidden');
            }
        });


        // Payload Modal
        document.getElementById('close-payload-modal-button').onclick = () => this.hideModal('payload-modal');
        document.getElementById('toggle-payload-schema-button').onclick = (e) => {
            const payloadContent = document.getElementById('payload-content');
            const schemaContent = document.getElementById('schema-content');
            const isSchemaVisible = !schemaContent.classList.contains('hidden');
            payloadContent.classList.toggle('hidden', !isSchemaVisible);
            schemaContent.classList.toggle('hidden', isSchemaVisible);
            e.target.textContent = isSchemaVisible ? 'Show Schema' : 'Show Payload';
        };


        // Templates Modal
        document.getElementById('templates-button').onclick = () => this.showTemplatesModal();
        document.getElementById('close-templates-modal-button').onclick = () => this.hideModal('templates-modal');
        document.getElementById('save-current-template-button').onclick = () => this.showSaveTemplateModal();

        // Save Template Modal
        document.getElementById('cancel-save-template-button').onclick = () => this.hideModal('save-template-modal');
        document.getElementById('confirm-save-template-button').onclick = () => this.saveTemplate();
        document.getElementById('delete-template-button').onclick = () => this.deleteTemplate();

        // Custom Param Modal
        document.getElementById('create-custom-param-button').onclick = () => this.showCustomParamModal();
        document.getElementById('cancel-custom-param-button').onclick = () => this.hideModal('custom-param-modal');
        document.getElementById('save-custom-param-button').onclick = () => this.saveCustomParameter();
        document.getElementById('manage-custom-params-button').onclick = () => this.showManageParamsModal();
        document.getElementById('close-manage-custom-params-button').onclick = () => this.hideModal('manage-custom-params-modal');
        
        // Logic Type Dropdown
        document.getElementById('logic-type-dropdown').onchange = (e) => this.toggleLogicSections(e.target.value);
        document.getElementById('add-mapping-button').onclick = () => this._addMappingRow();
        document.getElementById('import-script-button').onclick = () => document.getElementById('script-file-input').click();
        document.getElementById('script-file-input').onchange = (e) => this._handleScriptFileUpload(e);
        document.getElementById('scan-script-button').onclick = () => this._scanScriptForInputs();

        // Send Confirm Modal
        document.getElementById('cancel-send-button').onclick = () => this.hideModal('send-confirm-modal');
        document.getElementById('confirm-send-button').onclick = this.appContext.executeSend;

        // Recipient Modal
        document.getElementById('cancel-recipient-modal-button').onclick = () => this.hideModal('recipient-modal');
        document.getElementById('confirm-recipient-modal-button').onclick = () => this._confirmRecipientSelection();
        
        document.querySelectorAll('input[name="recipient-source"]').forEach(radio => {
            radio.addEventListener('change', (e) => this._handleRecipientSourceChange(e.target.value));
        });

        // Generic Confirm Modal
        document.getElementById('cancel-confirm-modal-button').onclick = () => this.hideModal('confirm-modal');
    }

    // --- Recipient Modal Logic ---
    async _fetchStudentCountForModal() {
        const statusEl = document.getElementById('recipient-modal-status');
        const confirmBtn = document.getElementById('confirm-recipient-modal-button');
        
        statusEl.textContent = 'Counting students...';
        statusEl.style.color = 'gray';
        confirmBtn.disabled = true;

        try {
            const selection = this._getRecipientSelectionFromModal();
            const students = await this.appContext.getStudentDataCore(selection);
            statusEl.textContent = `${students.length} student(s) will be selected.`;
            statusEl.style.color = 'green';
            confirmBtn.disabled = false;
            return students.length;
        } catch (error) {
            const message = error.userFacingMessage || (error.userFacing ? error.message : 'Error counting students.');
            statusEl.textContent = message;
            statusEl.style.color = 'red';
            return -1;
        }
    }
    
    _handleRecipientSourceChange(value) {
        document.getElementById('recipient-custom-sheet-container').classList.toggle('hidden', value !== 'custom');
        this._fetchStudentCountForModal();
    }
    
    _getRecipientSelectionFromModal() {
        const type = document.querySelector('input[name="recipient-source"]:checked').value;
        const customSheetName = document.getElementById('recipient-custom-sheet-name').value;
        const excludeDNC = document.getElementById('exclude-dnc-toggle').checked;
        const excludeFillColor = document.getElementById('exclude-fill-color-toggle').checked;
        return { type, customSheetName, excludeDNC, excludeFillColor };
    }
    
    async _confirmRecipientSelection() {
        const count = await this._fetchStudentCountForModal();
        if (count >= 0) {
            const selection = this._getRecipientSelectionFromModal();
            this.appContext.updateRecipientSelection(selection, count);
            this.hideModal('recipient-modal');
        }
    }

    showRecipientModal() {
        const { type, customSheetName, excludeDNC, excludeFillColor } = this.appContext.recipientSelection;
        
        document.getElementById(`recipient-${type}`).checked = true;
        document.getElementById('recipient-custom-sheet-name').value = customSheetName;
        document.getElementById('recipient-custom-sheet-container').classList.toggle('hidden', type !== 'custom');
        document.getElementById('exclude-dnc-toggle').checked = excludeDNC;
        document.getElementById('exclude-fill-color-toggle').checked = excludeFillColor;
        
        this.showModal('recipient-modal');
        this._fetchStudentCountForModal(); // Initial count on open
        
        // Add debounced fetching on custom sheet name input
        let debounceTimer;
        const customSheetInput = document.getElementById('recipient-custom-sheet-name');
        customSheetInput.oninput = () => {
            clearTimeout(debounceTimer);
            debounceTimer = setTimeout(() => this._fetchStudentCountForModal(), 500);
        };
        document.getElementById('exclude-dnc-toggle').onchange = () => this._fetchStudentCountForModal();
        document.getElementById('exclude-fill-color-toggle').onchange = () => this._fetchStudentCountForModal();
    }

    // --- Send Confirm Modal ---
    showSendConfirmModal() {
        const studentCount = this.appContext.studentDataCache.length;
        if (studentCount === 0) {
            this.appContext.getStudentDataWithUI().then(students => {
                if (students.length > 0) this.showSendConfirmModal();
            }).catch(() => {}); 
            return;
        }
        document.getElementById('send-confirm-message').textContent = `This will send a personalized email to ${studentCount} student(s). Are you sure you want to proceed?`;
        this.showModal('send-confirm-modal');
    }

    // --- Example Modal & Payload Modal ---
    async showExampleModal() {
        if (this.appContext.studentDataCache.length === 0) {
            try {
                await this.appContext.getStudentDataWithUI();
                if (this.appContext.studentDataCache.length === 0) {
                    const status = document.getElementById('status');
                    status.textContent = 'No students found to show an example.';
                    status.style.color = 'orange';
                    return;
                }
            } catch (error) {
                return; 
            }
        }
        this.showModal('example-modal');
        this.renderExampleForStudent(this.appContext.studentDataCache[this.currentExampleIndex]);
    }
    
    renderSearchResults(students) {
        const resultsContainer = document.getElementById('example-search-results');
        resultsContainer.innerHTML = '';
        if (students.length === 0) {
            resultsContainer.classList.add('hidden');
            return;
        }
        students.slice(0, 5).forEach(student => {
            const studentIndex = this.appContext.studentDataCache.findIndex(s => s.StudentEmail === student.StudentEmail);
            const div = document.createElement('div');
            div.className = 'px-3 py-2 hover:bg-gray-100 cursor-pointer text-sm';
            div.textContent = student.StudentName;
            div.onclick = () => {
                this.showAdjacentStudent(studentIndex);
                resultsContainer.classList.add('hidden');
                document.getElementById('example-search-input').value = '';
            };
            resultsContainer.appendChild(div);
        });
        resultsContainer.classList.remove('hidden');
    }

    renderExampleForStudent(student) {
        if (!student) return;
        const fromTemplate = this.appContext.fromParts.join('');
        const ccTemplate = this.appContext.ccRecipients;
        const subjectTemplate = document.getElementById('email-subject').value;
        const bodyTemplate = this.appContext.quill.root.innerHTML;

        document.getElementById('example-from').textContent = this.appContext.renderTemplate(fromTemplate, student);
        document.getElementById('example-to').textContent = student.StudentEmail || '(no email)';
        document.getElementById('example-cc').textContent = this.appContext.renderCCTemplate(ccTemplate, student);
        document.getElementById('example-subject').textContent = this.appContext.renderTemplate(subjectTemplate, student);
        document.getElementById('example-body').innerHTML = this.appContext.renderTemplate(bodyTemplate, student);
        
        const totalStudents = this.appContext.studentDataCache.length;
        document.getElementById('example-student-counter').textContent = `${this.currentExampleIndex + 1} of ${totalStudents}`;
        document.getElementById('prev-student-button').disabled = this.currentExampleIndex === 0;
        document.getElementById('next-student-button').disabled = this.currentExampleIndex >= totalStudents - 1;
    }
    
    showAdjacentStudent(index) {
        const total = this.appContext.studentDataCache.length;
        if (index >= 0 && index < total) {
            this.currentExampleIndex = index;
            this.renderExampleForStudent(this.appContext.studentDataCache[index]);
        }
    }
    
    showRandomStudent() {
        const total = this.appContext.studentDataCache.length;
        if (total > 1) {
            let randomIndex;
            do {
                randomIndex = Math.floor(Math.random() * total);
            } while (randomIndex === this.currentExampleIndex);
            this.showAdjacentStudent(randomIndex);
        }
    }

    showPayloadModal() {
        const fromTemplate = this.appContext.fromParts.join('');
        const subjectTemplate = document.getElementById('email-subject').value;
        const bodyTemplate = this.appContext.quill.root.innerHTML;

        this.currentPayload = this.appContext.studentDataCache.map(student => ({
            from: this.appContext.renderTemplate(fromTemplate, student),
            to: student.StudentEmail || '',
            cc: this.appContext.renderCCTemplate(this.appContext.ccRecipients, student),
            subject: this.appContext.renderTemplate(subjectTemplate, student),
            body: this.appContext.renderTemplate(bodyTemplate, student)
        })).filter(email => email.to && email.from);

        document.getElementById('payload-content').textContent = JSON.stringify(this.currentPayload, null, 2);
        document.getElementById('schema-content').textContent = JSON.stringify(this.currentSchema, null, 2);
        this.showModal('payload-modal');
    }

    // --- Template Modals ---
    async showTemplatesModal() {
        this.showModal('templates-modal');
        const templates = await this.appContext.getTemplates();
        const container = document.getElementById('templates-list-container');
        container.innerHTML = '';

        if (templates.length === 0) {
            container.innerHTML = '<p class="text-sm text-gray-500 text-center">No saved templates found.</p>';
            return;
        }

        // Group templates by folder
        const groupedTemplates = templates.reduce((acc, template) => {
            const folder = template.folder || 'Uncategorized';
            if (!acc[folder]) {
                acc[folder] = [];
            }
            acc[folder].push(template);
            return acc;
        }, {});

        // Render each folder
        Object.keys(groupedTemplates).sort().forEach(folderName => {
            const folderContainer = document.createElement('div');
            folderContainer.className = 'py-2';

            const folderHeader = document.createElement('div');
            folderHeader.className = 'flex items-center justify-between cursor-pointer px-1 py-1 rounded hover:bg-gray-100';
            folderHeader.innerHTML = `
                <div class="flex items-center">
                    <svg class="chevron-icon h-4 w-4 text-gray-500 mr-2" xmlns="http://www.w3.org/2000/svg" viewBox="0 0 20 20" fill="currentColor">
                        <path fill-rule="evenodd" d="M7.293 14.707a1 1 0 010-1.414L10.586 10 7.293 6.707a1 1 0 011.414-1.414l4 4a1 1 0 010 1.414l-4 4a1 1 0 01-1.414 0z" clip-rule="evenodd" />
                    </svg>
                    <h3 class="text-sm font-semibold text-gray-700">${folderName}</h3>
                </div>
                <span class="text-xs text-gray-500">${groupedTemplates[folderName].length}</span>
            `;

            const templatesList = document.createElement('div');
            templatesList.className = 'pl-4 mt-2 space-y-2 hidden'; // Initially hidden

            folderHeader.onclick = () => {
                templatesList.classList.toggle('hidden');
                folderHeader.querySelector('.chevron-icon').classList.toggle('chevron-open');
            };
            
            // Render templates within the folder
            groupedTemplates[folderName].forEach(template => {
                const div = document.createElement('div');
                div.className = 'p-3 border rounded-md hover:bg-gray-50';
                div.innerHTML = `
                    <div class="flex justify-between items-center">
                        <div>
                            <h4 class="font-semibold text-gray-800">${template.name}</h4>
                            <p class="text-xs text-gray-500">by ${template.author || 'Unknown'}</p>
                        </div>
                        <div class="flex gap-2">
                             <button data-template-id="${template.id}" class="apply-btn px-3 py-1 bg-blue-100 text-blue-800 text-xs font-semibold rounded-md hover:bg-blue-200">Apply</button>
                             <button data-template-id="${template.id}" class="edit-btn px-3 py-1 bg-gray-100 text-gray-800 text-xs font-semibold rounded-md hover:bg-gray-200">Edit</button>
                        </div>
                    </div>
                `;
                templatesList.appendChild(div);
            });
            
            folderContainer.appendChild(folderHeader);
            folderContainer.appendChild(templatesList);
            container.appendChild(folderContainer);
        });
        
        container.querySelectorAll('.apply-btn').forEach(btn => {
            btn.onclick = () => {
                const template = templates.find(t => t.id === btn.dataset.templateId);
                if(template) this._applyTemplate(template);
            };
        });
        container.querySelectorAll('.edit-btn').forEach(btn => {
            btn.onclick = () => {
                const template = templates.find(t => t.id === btn.dataset.templateId);
                if(template) this.showSaveTemplateModal(template);
            };
        });

    }

    _applyTemplate(template) {
        this.appContext.setFromParts(template.from || '');
        this.appContext.setCcRecipients(template.cc || []);
        document.getElementById('email-subject').value = template.subject || '';
        this.appContext.quill.root.innerHTML = template.body || '';
        this.hideModal('templates-modal');
    }

    showSaveTemplateModal(template = null) {
        this.hideModal('templates-modal');
        this.editingTemplate = template;
        
        document.getElementById('save-template-modal-title').textContent = template ? 'Edit Template' : 'Save Template';
        document.getElementById('template-name').value = template ? template.name : '';
        document.getElementById('template-folder').value = template ? template.folder : '';
        document.getElementById('template-author').value = template ? template.author : '';
        document.getElementById('delete-template-button').classList.toggle('hidden', !template);
        document.getElementById('save-template-status').textContent = '';
        this.showModal('save-template-modal');
    }

    async saveTemplate() {
        const name = document.getElementById('template-name').value.trim();
        const folder = document.getElementById('template-folder').value.trim();
        const author = document.getElementById('template-author').value.trim();
        const status = document.getElementById('save-template-status');

        if (!name) {
            status.textContent = "Template name is required.";
            status.style.color = 'red';
            return;
        }

        const currentTemplate = {
            id: this.editingTemplate ? this.editingTemplate.id : `template-${Date.now()}`,
            name,
            folder,
            author,
            from: this.appContext.fromParts,
            cc: this.appContext.ccRecipients,
            subject: document.getElementById('email-subject').value,
            body: this.appContext.quill.root.innerHTML
        };

        try {
            const templates = await this.appContext.getTemplates();
            const existingIndex = this.editingTemplate ? templates.findIndex(t => t.id === this.editingTemplate.id) : -1;

            if (existingIndex > -1) {
                templates[existingIndex] = currentTemplate;
            } else {
                templates.push(currentTemplate);
            }
            
            await this.appContext.saveTemplates(templates);
            this.hideModal('save-template-modal');
            this.editingTemplate = null;
        } catch (error) {
            status.textContent = "Error saving template.";
            status.style.color = 'red';
            console.error("Error saving template:", error);
        }
    }
    
    async deleteTemplate() {
        if (!this.editingTemplate) return;

        this.showConfirmModal({
            title: 'Delete Template',
            message: `Are you sure you want to delete the template "${this.editingTemplate.name}"?`,
            onConfirm: async () => {
                try {
                    let templates = await this.appContext.getTemplates();
                    templates = templates.filter(t => t.id !== this.editingTemplate.id);
                    await this.appContext.saveTemplates(templates);
                    this.hideModal('save-template-modal');
                    this.editingTemplate = null;
                } catch (error) {
                    const status = document.getElementById('save-template-status');
                    status.textContent = "Error deleting template.";
                    status.style.color = 'red';
                    console.error("Error deleting template:", error);
                }
            }
        });
    }

    // --- Custom Parameter Modals ---
    async showManageParamsModal() {
        this.hideModal('custom-param-modal');
        const params = await this.appContext.getCustomParameters();
        const container = document.getElementById('manage-custom-params-list');
        container.innerHTML = '';

        if (params.length === 0) {
            container.innerHTML = '<p class="text-sm text-gray-500 text-center">No custom parameters found.</p>';
        } else {
             params.forEach(param => {
                const div = document.createElement('div');
                div.className = 'flex justify-between items-center p-2 border-b';
                div.innerHTML = `
                    <div>
                        <p class="font-semibold text-gray-800">{${param.name}}</p>
                        <p class="text-xs text-gray-500">Source: ${param.sourceColumn}</p>
                    </div>
                    <div>
                        <button data-param-id="${param.id}" class="edit-param-btn px-3 py-1 bg-gray-100 text-gray-800 text-xs font-semibold rounded-md hover:bg-gray-200">Edit</button>
                        <button data-param-id="${param.id}" class="delete-param-btn px-3 py-1 bg-red-100 text-red-800 text-xs font-semibold rounded-md hover:bg-red-200">Delete</button>
                    </div>
                `;
                container.appendChild(div);
            });
            
            container.querySelectorAll('.edit-param-btn').forEach(btn => {
                btn.onclick = () => {
                    const paramId = btn.dataset.paramId;
                    const paramToEdit = params.find(p => p.id === paramId);
                    this.showCustomParamModal(paramToEdit);
                };
            });
             container.querySelectorAll('.delete-param-btn').forEach(btn => {
                btn.onclick = async () => {
                    const paramId = btn.dataset.paramId;
                    const paramToDelete = params.find(p => p.id === paramId);
                    if (!paramToDelete) return;

                    this.showConfirmModal({
                        title: 'Delete Parameter',
                        message: `Are you sure you want to delete the parameter "{${paramToDelete.name}}"? This may affect saved templates that use it.`,
                        onConfirm: async () => {
                            const updatedParams = params.filter(p => p.id !== paramId);
                            await this.appContext.saveCustomParameters(updatedParams);
                            await this.appContext.loadCustomParameters();
                            this.appContext.populateParameterButtons();
                            this.showManageParamsModal(); // Refresh the list
                        }
                    });
                };
            });
        }
        this.showModal('manage-custom-params-modal');
    }
    
    showCustomParamModal(param = null) {
        this.hideModal('manage-custom-params-modal');
        this.editingParam = param;
        document.getElementById('custom-param-modal-title').textContent = param ? 'Edit Custom Parameter' : 'Create Custom Parameter';
        document.getElementById('param-name').value = param ? param.name : '';
        document.getElementById('param-source-column').value = param ? param.sourceColumn : '';
        document.getElementById('save-param-status').textContent = '';

        // Reset logic sections
        const logicType = param ? param.logicType || '' : '';
        document.getElementById('logic-type-dropdown').value = logicType;
        this.toggleLogicSections(logicType);
        
        // Populate value mappings
        const mappingContainer = document.getElementById('param-mapping-container');
        mappingContainer.innerHTML = '';
        if (logicType === 'value-mapping' && param.mappings) {
            param.mappings.forEach(m => this._addMappingRow(m));
        }
        
        // Populate custom script
        const scriptEditor = document.getElementById('custom-script-editor');
        scriptEditor.value = '';
        if (logicType === 'custom-script' && param.script) {
            scriptEditor.value = param.script;
            if (param.scriptInputs) {
                this._renderScriptInputs(param.scriptInputs);
            }
        }
        this.showModal('custom-param-modal');
    }
    
    async saveCustomParameter() {
        const nameInput = document.getElementById('param-name');
        const sourceColumnInput = document.getElementById('param-source-column');
        const status = document.getElementById('save-param-status');

        const name = nameInput.value.trim().replace(/[^a-zA-Z0-9_]/g, '');
        const sourceColumn = sourceColumnInput.value.trim();

        if (!name || !sourceColumn) {
            status.textContent = 'Parameter Name and Source Column are required.';
            status.style.color = 'red';
            return;
        }
        nameInput.value = name;

        const logicType = document.getElementById('logic-type-dropdown').value;
        const newParam = {
            id: this.editingParam ? this.editingParam.id : `param-${Date.now()}`,
            name,
            sourceColumn,
            logicType: logicType || null
        };
        
        if (logicType === 'value-mapping') {
            newParam.mappings = this._getMappingsFromDOM();
        } else if (logicType === 'custom-script') {
            newParam.script = document.getElementById('custom-script-editor').value;
            newParam.scriptInputs = this._getScriptInputsFromDOM();
        }

        try {
            const params = await this.appContext.getCustomParameters();
            const existingIndex = this.editingParam ? params.findIndex(p => p.id === this.editingParam.id) : -1;

            if (existingIndex > -1) {
                params[existingIndex] = newParam;
            } else {
                 if (params.some(p => p.name.toLowerCase() === name.toLowerCase())) {
                    status.textContent = 'A parameter with this name already exists.';
                    status.style.color = 'red';
                    return;
                }
                params.push(newParam);
            }
            
            await this.appContext.saveCustomParameters(params);
            await this.appContext.loadCustomParameters();
            this.appContext.populateParameterButtons();
            this.hideModal('custom-param-modal');
            this.editingParam = null;

        } catch (error) {
            status.textContent = 'Error saving parameter.';
            status.style.color = 'red';
            console.error('Error saving custom parameter:', error);
        }
    }
    
    toggleLogicSections(selectedValue) {
        document.getElementById('value-mapping-logic-container').classList.toggle('hidden', selectedValue !== 'value-mapping');
        document.getElementById('custom-script-logic-container').classList.toggle('hidden', selectedValue !== 'custom-script');
    }

    _addMappingRow(mapping = { if: '', operator: 'eq', then: '' }) {
        const container = document.getElementById('param-mapping-container');
        const div = document.createElement('div');
        div.className = 'flex items-center gap-2 mapping-row';
        div.innerHTML = `
            <span class="text-sm">If value</span>
            <select class="operator-select w-28 border border-gray-300 rounded-md shadow-sm text-sm py-1 px-2">
                <option value="eq" ${mapping.operator === 'eq' ? 'selected' : ''}>is equal to</option>
                <option value="neq" ${mapping.operator === 'neq' ? 'selected' : ''}>is not equal to</option>
                <option value="contains" ${mapping.operator === 'contains' ? 'selected' : ''}>contains</option>
                <option value="does_not_contain" ${mapping.operator === 'does_not_contain' ? 'selected' : ''}>does not contain</option>
                <option value="starts_with" ${mapping.operator === 'starts_with' ? 'selected' : ''}>starts with</option>
                <option value="ends_with" ${mapping.operator === 'ends_with' ? 'selected' : ''}>ends with</option>
                <option value="gt" ${mapping.operator === 'gt' ? 'selected' : ''}>&gt;</option>
                <option value="lt" ${mapping.operator === 'lt' ? 'selected' : ''}>&lt;</option>
                <option value="gte" ${mapping.operator === 'gte' ? 'selected' : ''}>&gt;=</option>
                <option value="lte" ${mapping.operator === 'lte' ? 'selected' : ''}>&lt;=</option>
            </select>
            <input type="text" class="if-input flex-grow border border-gray-300 rounded-md shadow-sm text-sm py-1 px-2" value="${mapping.if}">
            <span class="text-sm">then</span>
            <input type="text" class="then-input flex-grow border border-gray-300 rounded-md shadow-sm text-sm py-1 px-2" value="${mapping.then}">
            <button class="remove-mapping-btn text-red-500 hover:text-red-700">&times;</button>
        `;
        div.querySelector('.remove-mapping-btn').onclick = () => div.remove();
        container.appendChild(div);
    }
    
    _getMappingsFromDOM() {
        const mappings = [];
        document.querySelectorAll('#param-mapping-container .mapping-row').forEach(row => {
            mappings.push({
                operator: row.querySelector('.operator-select').value,
                if: row.querySelector('.if-input').value,
                then: row.querySelector('.then-input').value
            });
        });
        return mappings;
    }

    _handleScriptFileUpload(event) {
        const file = event.target.files[0];
        if (file) {
            const reader = new FileReader();
            reader.onload = (e) => {
                document.getElementById('custom-script-editor').value = e.target.result;
            };
            reader.readAsText(file);
        }
    }

    _scanScriptForInputs() {
        const script = document.getElementById('custom-script-editor').value;
        // Regex to find variable declarations like "let myVar;" or "let myVar, myOtherVar;"
        const variableRegex = /let\s+([a-zA-Z0-9_,\s]+);/g;
        let match;
        const variables = new Set();
        while ((match = variableRegex.exec(script)) !== null) {
            match[1].split(',').forEach(v => {
                const trimmed = v.trim();
                if (trimmed) variables.add(trimmed);
            });
        }
        
        const existingInputs = this.editingParam ? this.editingParam.scriptInputs : {};
        const scriptInputs = {};
        variables.forEach(v => {
            scriptInputs[v] = existingInputs[v] || '';
        });
        
        this._renderScriptInputs(scriptInputs);
    }
    
    _renderScriptInputs(inputs) {
        const container = document.getElementById('script-variable-inputs-container');
        const list = document.getElementById('script-inputs-list');
        list.innerHTML = '';
        
        if (Object.keys(inputs).length === 0) {
            container.classList.add('hidden');
            return;
        }

        for (const varName in inputs) {
            const div = document.createElement('div');
            div.className = 'flex items-center gap-2 script-input-row';
            div.innerHTML = `
                <label class="w-32 text-sm text-gray-600 font-mono">${varName}</label>
                <input type="text" data-var-name="${varName}" class="flex-grow border border-gray-300 rounded-md shadow-sm text-sm py-1 px-2" value="${inputs[varName]}" placeholder="Source Column Name">
            `;
            list.appendChild(div);
        }
        container.classList.remove('hidden');
    }
    
    _getScriptInputsFromDOM() {
        const inputs = {};
        document.querySelectorAll('#script-inputs-list .script-input-row input').forEach(inputEl => {
            inputs[inputEl.dataset.varName] = inputEl.value;
        });
        return inputs;
    }

    // --- Generic Modal Visibility ---
    showModal(modalId) {
        document.getElementById(modalId).classList.remove('hidden');
    }

    hideModal(modalId) {
        document.getElementById(modalId).classList.add('hidden');
    }

    // --- Generic Confirmation Modal ---
    showConfirmModal({ title, message, onConfirm }) {
        document.getElementById('confirm-modal-title').textContent = title;
        document.getElementById('confirm-modal-message').textContent = message;

        const confirmButton = document.getElementById('confirm-confirm-modal-button');
        
        // Clone and replace to remove old listeners and ensure the correct onConfirm is used
        const newConfirmButton = confirmButton.cloneNode(true);
        confirmButton.parentNode.replaceChild(newConfirmButton, confirmButton);
        
        newConfirmButton.onclick = () => {
            this.hideModal('confirm-modal');
            onConfirm(); // Execute the specific callback for this confirmation
        };

        this.showModal('confirm-modal');
    }
}

