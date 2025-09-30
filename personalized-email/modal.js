// V-4.0 - 2025-09-30 - 1:04 PM EDT
export default class ModalManager {
    constructor(appContext) {
        this.appContext = appContext;
        this.editingTemplateId = null;
        this.editingParamName = null;
        this.currentScriptInputs = {};

        this._setupEventListeners();
    }

    _setupEventListeners() {
        // Example Modal
        document.getElementById('show-example-button').onclick = () => this.showExampleModal();
        document.getElementById('close-example-modal-button').onclick = () => this.hide('example-modal');

        // Payload Modal
        document.getElementById('close-payload-modal-button').onclick = () => this.hide('payload-modal');
        document.getElementById('toggle-payload-schema-button').onclick = () => this._togglePayloadSchema();

        // Templates Modal
        document.getElementById('templates-button').onclick = () => this.showTemplatesModal();
        document.getElementById('close-templates-modal-button').onclick = () => this.hide('templates-modal');
        document.getElementById('save-current-template-button').onclick = () => this.showSaveTemplateModal();
        
        // Save/Edit Template Modal
        document.getElementById('cancel-save-template-button').onclick = () => this.hide('save-template-modal');
        document.getElementById('confirm-save-template-button').onclick = () => this._saveTemplate();
        document.getElementById('delete-template-button').onclick = () => this._deleteTemplate();

        // Custom Parameter Modal
        document.getElementById('create-custom-param-button').onclick = () => this.showCustomParamModal();
        document.getElementById('cancel-custom-param-button').onclick = () => this.hide('custom-param-modal');
        document.getElementById('save-custom-param-button').onclick = () => this._saveCustomParameter();
        document.getElementById('add-mapping-button').onclick = () => this._addMappingRow();
        document.getElementById('logic-type-dropdown').onchange = (e) => this._toggleLogicContainers(e.target.value);
        
        // Manage Custom Parameters Modal
        document.getElementById('manage-custom-params-button').onclick = () => this.showManageCustomParamsModal();
        document.getElementById('close-manage-custom-params-button').onclick = () => this.hide('manage-custom-params-modal');
        
        // Send Confirmation Modal
        document.getElementById('cancel-send-button').onclick = () => this.hide('send-confirm-modal');
        document.getElementById('confirm-send-button').onclick = () => this.appContext.executeSend();

        // Custom Script Logic
        document.getElementById('scan-script-button').onclick = () => this._scanScriptForInputs();
        document.getElementById('import-script-button').onclick = () => document.getElementById('script-file-input').click();
        document.getElementById('script-file-input').onchange = (e) => this._handleScriptFileUpload(e);

    }

    show(modalId) {
        document.getElementById(modalId).classList.remove('hidden');
    }

    hide(modalId) {
        document.getElementById(modalId).classList.add('hidden');
    }

    // --- Example Modal Logic ---
    async showExampleModal() {
        try {
            const students = await this.appContext.getStudentData();
            if (students.length === 0) {
                document.getElementById('status').textContent = 'No students found to generate an example.';
                document.getElementById('status').style.color = 'orange';
                return;
            }
            const firstStudent = students[0];
            
            const from = this.appContext.renderTemplate(document.getElementById('email-from').value, firstStudent);
            const to = firstStudent.StudentEmail || '[No Email]';
            const cc = this.appContext.renderCCTemplate(this.appContext.ccRecipients, firstStudent);
            const subject = this.appContext.renderTemplate(document.getElementById('email-subject').value, firstStudent);
            const body = this.appContext.renderTemplate(this.appContext.quill.root.innerHTML, firstStudent);

            document.getElementById('example-from').textContent = from;
            document.getElementById('example-to').textContent = to;
            document.getElementById('example-cc').textContent = cc;
            document.getElementById('example-subject').textContent = subject;
            document.getElementById('example-body').innerHTML = body;

            this.show('example-modal');
        } catch (error) {
            console.error("Could not generate example:", error);
        }
    }

    // --- Payload Modal Logic ---
    async showPayloadModal() {
        try {
            const students = await this.appContext.getStudentData();
            if (students.length === 0) {
                document.getElementById('status').textContent = 'No students found to generate a payload.';
                document.getElementById('status').style.color = 'orange';
                return;
            }
            
            const fromTemplate = document.getElementById('email-from').value;
            const subjectTemplate = document.getElementById('email-subject').value;
            const bodyTemplate = this.appContext.quill.root.innerHTML;

            const payload = students.map(student => ({
                from: this.appContext.renderTemplate(fromTemplate, student),
                to: student.StudentEmail || '',
                cc: this.appContext.renderCCTemplate(this.appContext.ccRecipients, student),
                subject: this.appContext.renderTemplate(subjectTemplate, student),
                body: this.appContext.renderTemplate(bodyTemplate, student)
            })).filter(email => email.to && email.from);

            document.getElementById('payload-content').textContent = JSON.stringify(payload, null, 2);
            this.show('payload-modal');
        } catch (error) {
            console.error("Could not generate payload:", error);
        }
    }

    _togglePayloadSchema() {
        const payloadContent = document.getElementById('payload-content');
        const schemaContent = document.getElementById('schema-content');
        const button = document.getElementById('toggle-payload-schema-button');
        const title = document.getElementById('payload-modal-title');

        if (payloadContent.classList.contains('hidden')) {
            payloadContent.classList.remove('hidden');
            schemaContent.classList.add('hidden');
            button.textContent = 'Show Schema';
            title.textContent = 'Request Payload';
        } else {
            payloadContent.classList.add('hidden');
            schemaContent.classList.remove('hidden');
            button.textContent = 'Show Payload';
            title.textContent = 'Expected JSON Schema for Power Automate';
            
            if (!schemaContent.textContent) {
                 const schema = {
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
                schemaContent.textContent = JSON.stringify(schema, null, 2);
            }
        }
    }

    // --- Templates Modal Logic ---
    async showTemplatesModal() {
        await this._populateTemplatesList();
        this.show('templates-modal');
    }

    async _populateTemplatesList() {
        const container = document.getElementById('templates-list-container');
        container.innerHTML = '';
        const templates = await this.appContext.getTemplates();

        if (templates.length === 0) {
            container.innerHTML = '<p class="text-center text-gray-500 text-sm">No saved templates yet.</p>';
            return;
        }
        
        const groupedByAuthor = templates.reduce((acc, template) => {
            const author = template.author || 'Uncategorized';
            if (!acc[author]) {
                acc[author] = [];
            }
            acc[author].push(template);
            return acc;
        }, {});

        Object.keys(groupedByAuthor).sort().forEach(author => {
            const authorTemplates = groupedByAuthor[author];
            
            const authorContainer = document.createElement('div');
            
            const authorHeader = document.createElement('div');
            authorHeader.className = 'flex items-center justify-between p-2 rounded-md hover:bg-gray-100 cursor-pointer';
            
            const authorInfo = document.createElement('div');
            authorInfo.className = 'flex items-center space-x-2';
            
            const folderIcon = this._createIconSVG('folder');
            const authorName = document.createElement('span');
            authorName.className = 'font-semibold text-gray-700';
            authorName.textContent = author;
            
            authorInfo.append(folderIcon, authorName);

            const chevronIcon = this._createIconSVG('chevron');
            chevronIcon.classList.add('chevron-icon');

            authorHeader.append(authorInfo, chevronIcon);
            
            const templatesContainer = document.createElement('div');
            templatesContainer.className = 'hidden pl-6 border-l-2 border-gray-200 ml-2';

            authorTemplates.forEach(template => {
                const templateEl = this._createTemplateElement(template);
                templatesContainer.appendChild(templateEl);
            });

            authorHeader.onclick = () => {
                templatesContainer.classList.toggle('hidden');
                chevronIcon.classList.toggle('chevron-open');
            };

            authorContainer.append(authorHeader, templatesContainer);
            container.appendChild(authorContainer);
        });
    }

    _createTemplateElement(template) {
        const item = document.createElement('div');
        item.className = 'flex items-center justify-between p-2 my-1 rounded-md hover:bg-gray-50';

        const name = document.createElement('div');
        name.textContent = template.name;
        name.className = 'text-sm font-medium text-gray-800';
        
        const buttons = document.createElement('div');
        buttons.className = 'flex space-x-2';

        const loadButton = document.createElement('button');
        loadButton.textContent = 'Load';
        loadButton.className = 'px-2 py-1 bg-blue-100 text-blue-800 text-xs font-semibold rounded-md hover:bg-blue-200';
        loadButton.onclick = () => this._loadTemplate(template);
        
        const editButton = document.createElement('button');
        editButton.textContent = 'Edit';
        editButton.className = 'px-2 py-1 bg-gray-200 text-gray-800 text-xs font-semibold rounded-md hover:bg-gray-300';
        editButton.onclick = () => this.showSaveTemplateModal(template);

        buttons.append(loadButton, editButton);
        item.append(name, buttons);
        return item;
    }

    _loadTemplate(template) {
        document.getElementById('email-from').value = template.from;
        document.getElementById('email-subject').value = template.subject;
        this.appContext.quill.root.innerHTML = template.body;
        this.appContext.ccRecipients.splice(0, this.appContext.ccRecipients.length, ...template.cc);
        this.appContext.renderCCPills();
        this.hide('templates-modal');
    }

    async _deleteTemplate() {
        if (!this.editingTemplateId) return;

        let templates = await this.appContext.getTemplates();
        templates = templates.filter(t => t.id !== this.editingTemplateId);
        await this.appContext.saveTemplates(templates);

        this.hide('save-template-modal');
        await this._populateTemplatesList();
        this.show('templates-modal');
    }
    
    // --- Save/Edit Template Modal Logic ---
    showSaveTemplateModal(templateToEdit = null) {
        this.hide('templates-modal');
        const titleEl = document.getElementById('save-template-modal-title');
        const deleteBtn = document.getElementById('delete-template-button');
        
        if (templateToEdit) {
            this.editingTemplateId = templateToEdit.id;
            titleEl.textContent = 'Edit Template';
            document.getElementById('template-name').value = templateToEdit.name;
            document.getElementById('template-author').value = templateToEdit.author;
            deleteBtn.classList.remove('hidden');
        } else {
            this.editingTemplateId = null;
            titleEl.textContent = 'Save New Template';
            document.getElementById('template-name').value = '';
            document.getElementById('template-author').value = '';
            deleteBtn.classList.add('hidden');
        }
        
        document.getElementById('save-template-status').textContent = '';
        this.show('save-template-modal');
    }

    async _saveTemplate() {
        const name = document.getElementById('template-name').value.trim();
        const author = document.getElementById('template-author').value.trim();
        const status = document.getElementById('save-template-status');

        if (!name || !author) {
            status.textContent = 'Name and Author are required.';
            status.style.color = 'red';
            return;
        }

        const templates = await this.appContext.getTemplates();

        if (this.editingTemplateId) {
            // Update existing template
            const templateIndex = templates.findIndex(t => t.id === this.editingTemplateId);
            if (templateIndex > -1) {
                const updatedTemplate = {
                    ...templates[templateIndex],
                    name,
                    author,
                    from: document.getElementById('email-from').value,
                    subject: document.getElementById('email-subject').value,
                    body: this.appContext.quill.root.innerHTML,
                    cc: [...this.appContext.ccRecipients],
                };
                templates[templateIndex] = updatedTemplate;
            }
        } else {
            // Create new template
            const newTemplate = {
                id: 'tpl-' + Math.random().toString(36).substr(2, 9),
                name,
                author,
                from: document.getElementById('email-from').value,
                subject: document.getElementById('email-subject').value,
                body: this.appContext.quill.root.innerHTML,
                cc: [...this.appContext.ccRecipients],
                createdAt: new Date().toISOString()
            };
            templates.push(newTemplate);
        }
        
        await this.appContext.saveTemplates(templates);

        status.textContent = 'Template saved successfully!';
        status.style.color = 'green';

        setTimeout(() => {
            this.hide('save-template-modal');
            this.showTemplatesModal();
        }, 1500);
    }
    
    // --- Custom Parameter Modal Logic ---
    showCustomParamModal(paramToEdit = null) {
        this._resetCustomParamModal();
        if (paramToEdit) {
            this.editingParamName = paramToEdit.name;
            document.getElementById('custom-param-modal-title').textContent = 'Edit Custom Parameter';
            document.getElementById('param-name').value = paramToEdit.name;
            document.getElementById('param-source-column').value = paramToEdit.sourceColumn;
            
            const logicType = paramToEdit.logicType || '';
            document.getElementById('logic-type-dropdown').value = logicType;
            this._toggleLogicContainers(logicType);

            if (logicType === 'value-mapping' && paramToEdit.mappings) {
                paramToEdit.mappings.forEach(m => this._addMappingRow(m));
            }
            if (logicType === 'custom-script' && paramToEdit.script) {
                document.getElementById('custom-script-editor').value = paramToEdit.script;
                if (paramToEdit.scriptInputs) {
                    this.currentScriptInputs = { ...paramToEdit.scriptInputs };
                    this._renderScriptInputFields();
                }
            }
        }
        this.show('custom-param-modal');
    }

    _resetCustomParamModal() {
        this.editingParamName = null;
        this.currentScriptInputs = {};
        document.getElementById('custom-param-modal-title').textContent = 'Create Custom Parameter';
        document.getElementById('param-name').value = '';
        document.getElementById('param-source-column').value = '';
        document.getElementById('param-mapping-container').innerHTML = '';
        document.getElementById('save-param-status').textContent = '';
        document.getElementById('logic-type-dropdown').value = '';
        document.getElementById('custom-script-editor').value = '';
        this._renderScriptInputFields();
        this._toggleLogicContainers('');
    }

    async _saveCustomParameter() {
        const name = document.getElementById('param-name').value.trim();
        const sourceColumn = document.getElementById('param-source-column').value.trim();
        const logicType = document.getElementById('logic-type-dropdown').value;
        const status = document.getElementById('save-param-status');

        if (!name || !sourceColumn) {
            status.textContent = 'Parameter Name and Source Column are required.';
            status.style.color = 'red';
            return;
        }
        if (!/^[a-zA-Z0-9_]+$/.test(name)) {
            status.textContent = 'Parameter Name can only contain letters, numbers, and underscores.';
            status.style.color = 'red';
            return;
        }

        const newParam = { name, sourceColumn, logicType };

        if (logicType === 'value-mapping') {
            newParam.mappings = this._getMappingsFromDOM();
        }
        if (logicType === 'custom-script') {
            newParam.script = document.getElementById('custom-script-editor').value.trim();
            this._updateScriptInputsFromDOM();
            newParam.scriptInputs = this.currentScriptInputs;
        }

        let params = await this.appContext.getCustomParameters();
        if (this.editingParamName && this.editingParamName !== name) {
            // If name changed, ensure new name isn't a duplicate
            if (params.some(p => p.name === name)) {
                status.textContent = 'A parameter with this name already exists.';
                status.style.color = 'red';
                return;
            }
            // Remove old param
            params = params.filter(p => p.name !== this.editingParamName);
        } else if (!this.editingParamName) {
            // If creating new, check for duplicate name
             if (params.some(p => p.name === name)) {
                status.textContent = 'A parameter with this name already exists.';
                status.style.color = 'red';
                return;
            }
        }
        
        const existingIndex = params.findIndex(p => p.name === name);
        if (existingIndex > -1) {
            params[existingIndex] = newParam;
        } else {
            params.push(newParam);
        }

        await this.appContext.saveCustomParameters(params);
        await this.appContext.loadCustomParameters();
        this.appContext.populateParameterButtons();
        
        status.textContent = 'Parameter saved!';
        status.style.color = 'green';

        setTimeout(() => {
            this.hide('custom-param-modal');
            this.showManageCustomParamsModal(); // Show updated list
        }, 1000);
    }
    
    // --- Manage Custom Parameters Modal Logic ---
    async showManageCustomParamsModal() {
        this.hide('custom-param-modal');
        await this._populateManageParamsList();
        this.show('manage-custom-params-modal');
    }

    async _populateManageParamsList() {
        const container = document.getElementById('manage-custom-params-list');
        container.innerHTML = '';
        const params = await this.appContext.getCustomParameters();

        if (params.length === 0) {
            container.innerHTML = '<p class="text-center text-gray-500 text-sm">No custom parameters created yet.</p>';
            return;
        }
        
        params.forEach(param => {
            const item = document.createElement('div');
            item.className = 'flex items-center justify-between p-2 my-1 rounded-md hover:bg-gray-50';

            const info = document.createElement('div');
            const name = document.createElement('span');
            name.textContent = `{${param.name}}`;
            name.className = 'text-sm font-medium text-gray-800';
            const source = document.createElement('span');
            source.textContent = ` (from: ${param.sourceColumn})`;
            source.className = 'text-xs text-gray-500';
            info.append(name, source);
            
            const buttons = document.createElement('div');
            buttons.className = 'flex space-x-2';

            const editButton = document.createElement('button');
            editButton.textContent = 'Edit';
            editButton.className = 'px-2 py-1 bg-blue-100 text-blue-800 text-xs font-semibold rounded-md hover:bg-blue-200';
            editButton.onclick = () => {
                this.hide('manage-custom-params-modal');
                this.showCustomParamModal(param);
            };
            
            const deleteButton = document.createElement('button');
            deleteButton.textContent = 'Delete';
            deleteButton.className = 'px-2 py-1 bg-red-100 text-red-800 text-xs font-semibold rounded-md hover:bg-red-200';
            deleteButton.onclick = () => this._deleteCustomParameter(param.name);
            
            buttons.append(editButton, deleteButton);
            item.append(info, buttons);
            container.appendChild(item);
        });
    }

    async _deleteCustomParameter(paramName) {
        let params = await this.appContext.getCustomParameters();
        params = params.filter(p => p.name !== paramName);
        await this.appContext.saveCustomParameters(params);
        await this.appContext.loadCustomParameters();
        this.appContext.populateParameterButtons();
        await this._populateManageParamsList();
    }
    
    // --- Custom Parameter - Value Mapping UI Logic ---
    _toggleLogicContainers(selectedValue) {
        const mappingContainer = document.getElementById('value-mapping-logic-container');
        const scriptContainer = document.getElementById('custom-script-logic-container');
        mappingContainer.classList.toggle('hidden', selectedValue !== 'value-mapping');
        scriptContainer.classList.toggle('hidden', selectedValue !== 'custom-script');
    }
    
    _addMappingRow(mapping = { if: '', operator: 'eq', then: '' }) {
        const container = document.getElementById('param-mapping-container');
        const row = document.createElement('div');
        row.className = 'flex items-center gap-2 mapping-row';
        row.innerHTML = `
            <span class="text-sm">If cell</span>
            <select class="operator-select w-32 px-2 py-1 border border-gray-300 rounded-md text-sm">
                <option value="eq" ${mapping.operator === 'eq' ? 'selected' : ''}>is equal to</option>
                <option value="neq" ${mapping.operator === 'neq' ? 'selected' : ''}>is not equal to</option>
                <option value="contains" ${mapping.operator === 'contains' ? 'selected' : ''}>contains</option>
                <option value="does_not_contain" ${mapping.operator === 'does_not_contain' ? 'selected' : ''}>does not contain</option>
                <option value="starts_with" ${mapping.operator === 'starts_with' ? 'selected' : ''}>starts with</option>
                <option value="ends_with" ${mapping.operator === 'ends_with' ? 'selected' : ''}>ends with</option>
                <option value="gt" ${mapping.operator === 'gt' ? 'selected' : ''}>&gt; (number)</option>
                <option value="lt" ${mapping.operator === 'lt' ? 'selected' : ''}>&lt; (number)</option>
                <option value="gte" ${mapping.operator === 'gte' ? 'selected' : ''}>&gt;= (number)</option>
                <option value="lte" ${mapping.operator === 'lte' ? 'selected' : ''}>&lt;= (number)</option>
            </select>
            <input type="text" class="if-input flex-grow px-2 py-1 border border-gray-300 rounded-md text-sm" placeholder="Value..." value="${mapping.if}">
            <span class="text-sm">then</span>
            <input type="text" class="then-input flex-grow px-2 py-1 border border-gray-300 rounded-md text-sm" placeholder="Result..." value="${mapping.then}">
            <button class="remove-mapping-btn text-red-500 hover:text-red-700">&times;</button>
        `;
        row.querySelector('.remove-mapping-btn').onclick = () => row.remove();
        container.appendChild(row);
    }
    
    _getMappingsFromDOM() {
        const mappings = [];
        document.querySelectorAll('.mapping-row').forEach(row => {
            const operator = row.querySelector('.operator-select').value;
            const ifValue = row.querySelector('.if-input').value;
            const thenValue = row.querySelector('.then-input').value;
            if (ifValue) {
                mappings.push({ if: ifValue, operator, then: thenValue });
            }
        });
        return mappings;
    }

    // --- Custom Parameter - Custom Script Logic ---
     _scanScriptForInputs() {
        const script = document.getElementById('custom-script-editor').value;
        // Regex to find variable declarations (let, const, var) that are NOT assigned a value
        const regex = /\b(?:let|const|var)\s+([a-zA-Z_$][a-zA-Z0-9_$]*)(?!\s*=)/g;
        let match;
        const newInputs = new Set();
        while ((match = regex.exec(script)) !== null) {
            const varName = match[1];
            // Exclude standard provided arguments
            if (varName !== 'getWorksheet' && varName !== 'sourceColumnValue') {
                newInputs.add(varName);
            }
        }

        // Preserve existing values if the variable still exists
        const updatedScriptInputs = {};
        newInputs.forEach(name => {
            updatedScriptInputs[name] = this.currentScriptInputs[name] || '';
        });
        this.currentScriptInputs = updatedScriptInputs;
        
        this._renderScriptInputFields();
    }

    _renderScriptInputFields() {
        const container = document.getElementById('script-inputs-list');
        const parentContainer = document.getElementById('script-variable-inputs-container');
        container.innerHTML = '';

        const inputNames = Object.keys(this.currentScriptInputs);

        if (inputNames.length === 0) {
            parentContainer.classList.add('hidden');
            return;
        }
        
        inputNames.forEach(varName => {
            const row = document.createElement('div');
            row.className = 'flex items-center gap-2';
            row.innerHTML = `
                <label for="script-input-${varName}" class="w-1/3 text-sm font-mono text-gray-600">${varName}</label>
                <input type="text" id="script-input-${varName}" data-varname="${varName}" class="script-input-field flex-grow px-2 py-1 border border-gray-300 rounded-md text-sm" placeholder="Source Column Name" value="${this.currentScriptInputs[varName] || ''}">
            `;
            container.appendChild(row);
        });

        parentContainer.classList.remove('hidden');
    }

    _updateScriptInputsFromDOM() {
        document.querySelectorAll('.script-input-field').forEach(input => {
            const varName = input.dataset.varname;
            this.currentScriptInputs[varName] = input.value.trim();
        });
    }

    _handleScriptFileUpload(event) {
        const file = event.target.files[0];
        if (file) {
            const reader = new FileReader();
            reader.onload = (e) => {
                document.getElementById('custom-script-editor').value = e.target.result;
                this._scanScriptForInputs();
            };
            reader.readAsText(file);
        }
        // Reset file input so the same file can be loaded again
        event.target.value = '';
    }

    // --- Send Confirmation Modal Logic ---
    async showSendConfirmModal() {
        try {
            await this.appContext.getStudentData();
            const count = this.appContext.studentDataCache.length;
            if (count === 0) {
                document.getElementById('status').textContent = 'No students found to send emails to.';
                document.getElementById('status').style.color = 'orange';
                return;
            }
            const message = `You are about to send ${count} ${count === 1 ? 'email' : 'emails'}. Do you want to proceed?`;
            document.getElementById('send-confirm-message').textContent = message;
            this.show('send-confirm-modal');
        } catch (error) {
            // Error is handled in getStudentData, no need to show another message.
        }
    }
    
    // --- Helper to create SVG icons ---
    _createIconSVG(type) {
        const svg = document.createElementNS('http://www.w3.org/2000/svg', 'svg');
        svg.setAttribute('class', 'h-5 w-5 text-gray-500');
        svg.setAttribute('viewBox', '0 0 20 20');
        svg.setAttribute('fill', 'currentColor');
        let pathData = '';
        switch(type) {
            case 'folder':
                pathData = 'M2 6a2 2 0 012-2h5l2 2h5a2 2 0 012 2v6a2 2 0 01-2 2H4a2 2 0 01-2-2V6z';
                break;
            case 'chevron':
                pathData = 'M7.293 14.707a1 1 0 010-1.414L10.586 10 7.293 6.707a1 1 0 011.414-1.414l4 4a1 1 0 010 1.414l-4 4a1 1 0 01-1.414 0z';
                break;
        }
        const path = document.createElementNS('http://www.w3.org/2000/svg', 'path');
        path.setAttribute('d', pathData);
        svg.appendChild(path);
        return svg;
    }
}

