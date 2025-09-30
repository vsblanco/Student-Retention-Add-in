// V-4.7 - 2025-09-30 - 2:50 PM EDT
export default class ModalManager {
    constructor(appContext) {
        this.appContext = appContext;
        this.editingTemplateId = null;
        this.editingParamName = null;
        this.currentScriptInputs = {};
        this.currentExampleIndex = 0;
        this.studentsForExample = [];
        this.currentRecipientSelection = {};
        this.tempStudentCount = 0;

        this._setupEventListeners();
    }

    _setupEventListeners() {
        // Example Modal
        document.getElementById('show-example-button').onclick = () => this.showExampleModal();
        document.getElementById('close-example-modal-button').onclick = () => this.hide('example-modal');
        document.getElementById('prev-student-button').onclick = () => this._navigateExample(-1);
        document.getElementById('next-student-button').onclick = () => this._navigateExample(1);
        document.getElementById('random-student-button').onclick = () => this._randomizeExample();
        document.getElementById('search-student-button').onclick = () => this._toggleExampleSearch();
        document.getElementById('example-search-input').oninput = (e) => this._filterStudents(e.target.value);

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

        // Recipient Modal
        document.getElementById('cancel-recipient-modal-button').onclick = () => this.hide('recipient-modal');
        document.getElementById('confirm-recipient-modal-button').onclick = () => this._confirmRecipientSelection();
        document.querySelectorAll('input[name="recipient-source"]').forEach(radio => {
            radio.onchange = () => this._handleRecipientSourceChange();
        });
        document.getElementById('recipient-custom-sheet-name').oninput = () => this._handleRecipientSourceChange();
        document.getElementById('exclude-dnc-toggle').onchange = () => this._handleRecipientSourceChange();
    }

    show(modalId) {
        document.getElementById(modalId).classList.remove('hidden');
    }

    hide(modalId) {
        document.getElementById(modalId).classList.add('hidden');
    }

    // --- Recipient Modal Logic ---
    showRecipientModal() {
        this.currentRecipientSelection = { ...this.appContext.recipientSelection };
        this.tempStudentCount = 0;

        const { type, customSheetName, excludeDNC } = this.currentRecipientSelection;
        document.querySelector(`input[name="recipient-source"][value="${type}"]`).checked = true;
        document.getElementById('recipient-custom-sheet-name').value = customSheetName || '';
        document.getElementById('exclude-dnc-toggle').checked = excludeDNC;


        this._updateRecipientModalUI();
        this._fetchStudentCountForModal();
        this.show('recipient-modal');
    }

    _handleRecipientSourceChange() {
        const selectedType = document.querySelector('input[name="recipient-source"]:checked').value;
        const customSheetName = document.getElementById('recipient-custom-sheet-name').value.trim();
        const excludeDNC = document.getElementById('exclude-dnc-toggle').checked;
        this.currentRecipientSelection = { type: selectedType, customSheetName, excludeDNC };
        this._updateRecipientModalUI();
        this._fetchStudentCountForModal();
    }

    _updateRecipientModalUI() {
        const isCustom = this.currentRecipientSelection.type === 'custom';
        document.getElementById('recipient-custom-sheet-container').classList.toggle('hidden', !isCustom);
    }

    async _fetchStudentCountForModal() {
        const statusEl = document.getElementById('recipient-modal-status');
        const confirmBtn = document.getElementById('confirm-recipient-modal-button');
        
        statusEl.textContent = 'Counting students...';
        confirmBtn.disabled = true;

        try {
            const students = await this.appContext.getStudentDataCore(this.currentRecipientSelection);
            this.tempStudentCount = students.length;
            statusEl.textContent = `${this.tempStudentCount} student${this.tempStudentCount !== 1 ? 's' : ''} found.`;
            confirmBtn.disabled = false; // Always allow confirmation, even for 0 students
        } catch (error) {
            this.tempStudentCount = 0;
            statusEl.textContent = error.userFacingMessage || (error.userFacing ? error.message : 'An error occurred.');
        }
    }

    _confirmRecipientSelection() {
        this.appContext.updateRecipientSelection(this.currentRecipientSelection, this.tempStudentCount);
        this.hide('recipient-modal');
        // Trigger a data fetch with UI updates for the main screen
        this.appContext.getStudentDataWithUI().catch(() => {
            // Reset button if the fetch fails after confirmation
            this.appContext.updateRecipientSelection(this.appContext.recipientSelection, -1);
        });
    }

    // --- Example Modal Logic ---
    async showExampleModal() {
        if (this.appContext.studentDataCache.length === 0) {
             document.getElementById('status').textContent = 'Please select recipients before viewing an example.';
             document.getElementById('status').style.color = 'orange';
             return;
        }
        this.studentsForExample = this.appContext.studentDataCache;
        this.currentExampleIndex = 0;
        this._resetExampleSearch();
        this._renderExampleForIndex(this.currentExampleIndex);
        this.show('example-modal');
    }

    _renderExampleForIndex(index) {
        const students = this.studentsForExample;
        if (!students || students.length === 0 || index < 0 || index >= students.length) return;
        
        const student = students[index];
        
        const from = this.appContext.renderTemplate(document.getElementById('email-from').value, student);
        const to = student.StudentEmail || '[No Email]';
        const cc = this.appContext.renderCCTemplate(this.appContext.ccRecipients, student);
        const subject = this.appContext.renderTemplate(document.getElementById('email-subject').value, student);
        const body = this.appContext.renderTemplate(this.appContext.quill.root.innerHTML, student);

        document.getElementById('example-from').textContent = from;
        document.getElementById('example-to').textContent = to;
        document.getElementById('example-cc').textContent = cc;
        document.getElementById('example-subject').textContent = subject;
        document.getElementById('example-body').innerHTML = body;

        document.getElementById('example-student-counter').textContent = `Student: ${index + 1} / ${students.length}`;
        document.getElementById('prev-student-button').disabled = index === 0;
        document.getElementById('next-student-button').disabled = index === students.length - 1;
    }

    _navigateExample(direction) {
        const newIndex = this.currentExampleIndex + direction;
        if (newIndex >= 0 && newIndex < this.studentsForExample.length) {
            this.currentExampleIndex = newIndex;
            this._renderExampleForIndex(this.currentExampleIndex);
            this._resetExampleSearch();
        }
    }

    _randomizeExample() {
        if (this.studentsForExample.length > 0) {
            const randomIndex = Math.floor(Math.random() * this.studentsForExample.length);
            this.currentExampleIndex = randomIndex;
            this._renderExampleForIndex(this.currentExampleIndex);
            this._resetExampleSearch();
        }
    }

    _toggleExampleSearch() {
        document.getElementById('example-search-container').classList.toggle('hidden');
    }

    _resetExampleSearch() {
        document.getElementById('example-search-container').classList.add('hidden');
        document.getElementById('example-search-input').value = '';
        document.getElementById('example-search-results').classList.add('hidden');
        document.getElementById('example-search-results').innerHTML = '';
    }
    
    _filterStudents(searchTerm) {
        const resultsContainer = document.getElementById('example-search-results');
        const term = searchTerm.toLowerCase().trim();

        if (term.length === 0) {
            resultsContainer.innerHTML = '';
            resultsContainer.classList.add('hidden');
            return;
        }

        const matches = this.studentsForExample.map((student, index) => ({ student, originalIndex: index }))
            .filter(item => item.student.StudentName && item.student.StudentName.toLowerCase().includes(term));
        this._renderSearchResults(matches);
    }

    _renderSearchResults(matches) {
        const resultsContainer = document.getElementById('example-search-results');
        resultsContainer.innerHTML = '';

        if (matches.length === 0) {
            resultsContainer.innerHTML = '<div class="px-3 py-2 text-sm text-gray-500">No matches found.</div>';
        } else {
            matches.slice(0, 10).forEach(match => {
                const item = document.createElement('div');
                item.className = 'px-3 py-2 text-sm text-gray-700 cursor-pointer hover:bg-gray-100';
                item.textContent = match.student.StudentName;
                item.onclick = () => this._selectSearchResult(match.originalIndex);
                resultsContainer.appendChild(item);
            });
        }
        resultsContainer.classList.remove('hidden');
    }

    _selectSearchResult(originalIndex) {
        this.currentExampleIndex = originalIndex;
        this._renderExampleForIndex(this.currentExampleIndex);
        this._resetExampleSearch();
    }

    async showPayloadModal() {
        if (this.appContext.studentDataCache.length === 0) {
             document.getElementById('status').textContent = 'Please select recipients before viewing payload.';
             document.getElementById('status').style.color = 'orange';
             return;
        }
        
        const fromTemplate = document.getElementById('email-from').value;
        const subjectTemplate = document.getElementById('email-subject').value;
        const bodyTemplate = this.appContext.quill.root.innerHTML;

        const payload = this.appContext.studentDataCache.map(student => ({
            from: this.appContext.renderTemplate(fromTemplate, student),
            to: student.StudentEmail || '',
            cc: this.appContext.renderCCTemplate(this.appContext.ccRecipients, student),
            subject: this.appContext.renderTemplate(subjectTemplate, student),
            body: this.appContext.renderTemplate(bodyTemplate, student)
        })).filter(email => email.to && email.from);

        document.getElementById('payload-content').textContent = JSON.stringify(payload, null, 2);
        this.show('payload-modal');
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
                 schemaContent.textContent = JSON.stringify({
                    "type": "array",
                    "items": {
                        "type": "object",
                        "properties": { "from": { "type": "string" }, "to": { "type": "string" }, "cc": { "type": "string" }, "subject": { "type": "string" }, "body": { "type": "string" } },
                        "required": ["from", "to", "subject", "body"]
                    }
                }, null, 2);
            }
        }
    }

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
            if (!acc[author]) acc[author] = [];
            acc[author].push(template);
            return acc;
        }, {});

        Object.keys(groupedByAuthor).sort().forEach(author => {
            const authorHeader = document.createElement('div');
            authorHeader.className = 'flex items-center justify-between p-2 rounded-md hover:bg-gray-100 cursor-pointer';
            authorHeader.innerHTML = `
                <div class="flex items-center space-x-2">
                    <svg class="h-5 w-5 text-gray-500" viewBox="0 0 20 20" fill="currentColor"><path d="M2 6a2 2 0 012-2h5l2 2h5a2 2 0 012 2v6a2 2 0 01-2 2H4a2 2 0 01-2-2V6z"></path></svg>
                    <span class="font-semibold text-gray-700">${author}</span>
                </div>
                <svg class="h-5 w-5 text-gray-500 chevron-icon" viewBox="0 0 20 20" fill="currentColor"><path d="M7.293 14.707a1 1 0 010-1.414L10.586 10 7.293 6.707a1 1 0 011.414-1.414l4 4a1 1 0 010 1.414l-4 4a1 1 0 01-1.414 0z"></path></svg>
            `;
            const templatesContainer = document.createElement('div');
            templatesContainer.className = 'hidden pl-6 border-l-2 border-gray-200 ml-2';
            groupedByAuthor[author].forEach(template => templatesContainer.appendChild(this._createTemplateElement(template)));
            
            authorHeader.onclick = () => {
                templatesContainer.classList.toggle('hidden');
                authorHeader.querySelector('.chevron-icon').classList.toggle('chevron-open');
            };
            const authorContainer = document.createElement('div');
            authorContainer.append(authorHeader, templatesContainer);
            container.appendChild(authorContainer);
        });
    }

    _createTemplateElement(template) {
        const item = document.createElement('div');
        item.className = 'flex items-center justify-between p-2 my-1 rounded-md hover:bg-gray-50';
        item.innerHTML = `
            <div class="text-sm font-medium text-gray-800">${template.name}</div>
            <div class="flex space-x-2">
                <button class="load-btn px-2 py-1 bg-blue-100 text-blue-800 text-xs font-semibold rounded-md hover:bg-blue-200">Load</button>
                <button class="edit-btn px-2 py-1 bg-gray-200 text-gray-800 text-xs font-semibold rounded-md hover:bg-gray-300">Edit</button>
            </div>
        `;
        item.querySelector('.load-btn').onclick = () => this._loadTemplate(template);
        item.querySelector('.edit-btn').onclick = () => this.showSaveTemplateModal(template);
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
        const newTemplateData = {
            name, author,
            from: document.getElementById('email-from').value,
            subject: document.getElementById('email-subject').value,
            body: this.appContext.quill.root.innerHTML,
            cc: [...this.appContext.ccRecipients],
        };

        if (this.editingTemplateId) {
            const templateIndex = templates.findIndex(t => t.id === this.editingTemplateId);
            if (templateIndex > -1) templates[templateIndex] = { ...templates[templateIndex], ...newTemplateData };
        } else {
            templates.push({ ...newTemplateData, id: 'tpl-' + Date.now(), createdAt: new Date().toISOString() });
        }
        
        await this.appContext.saveTemplates(templates);
        status.textContent = 'Template saved successfully!';
        status.style.color = 'green';

        setTimeout(() => {
            this.hide('save-template-modal');
            this.showTemplatesModal();
        }, 1500);
    }
    
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

            if (logicType === 'value-mapping' && paramToEdit.mappings) paramToEdit.mappings.forEach(m => this._addMappingRow(m));
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
        if (logicType === 'value-mapping') newParam.mappings = this._getMappingsFromDOM();
        if (logicType === 'custom-script') {
            newParam.script = document.getElementById('custom-script-editor').value.trim();
            this._updateScriptInputsFromDOM();
            newParam.scriptInputs = this.currentScriptInputs;
        }

        let params = await this.appContext.getCustomParameters();
        if (this.editingParamName && this.editingParamName !== name) {
            if (params.some(p => p.name === name)) {
                status.textContent = 'A parameter with this name already exists.';
                status.style.color = 'red';
                return;
            }
            params = params.filter(p => p.name !== this.editingParamName);
        } else if (!this.editingParamName && params.some(p => p.name === name)) {
            status.textContent = 'A parameter with this name already exists.';
            status.style.color = 'red';
            return;
        }
        
        const existingIndex = params.findIndex(p => p.name === name);
        if (existingIndex > -1) params[existingIndex] = newParam;
        else params.push(newParam);

        await this.appContext.saveCustomParameters(params);
        await this.appContext.loadCustomParameters();
        this.appContext.populateParameterButtons();
        status.textContent = 'Parameter saved!';
        status.style.color = 'green';
        setTimeout(() => { this.hide('custom-param-modal'); this.showManageCustomParamsModal(); }, 1000);
    }
    
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
            item.innerHTML = `
                <div>
                    <span class="text-sm font-medium text-gray-800">{${param.name}}</span>
                    <span class="text-xs text-gray-500">(from: ${param.sourceColumn})</span>
                </div>
                <div class="flex space-x-2">
                    <button class="edit-btn px-2 py-1 bg-blue-100 text-blue-800 text-xs font-semibold rounded-md hover:bg-blue-200">Edit</button>
                    <button class="delete-btn px-2 py-1 bg-red-100 text-red-800 text-xs font-semibold rounded-md hover:bg-red-200">Delete</button>
                </div>
            `;
            item.querySelector('.edit-btn').onclick = () => { this.hide('manage-custom-params-modal'); this.showCustomParamModal(param); };
            item.querySelector('.delete-btn').onclick = () => this._deleteCustomParameter(param.name);
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
    
    _toggleLogicContainers(selectedValue) {
        document.getElementById('value-mapping-logic-container').classList.toggle('hidden', selectedValue !== 'value-mapping');
        document.getElementById('custom-script-logic-container').classList.toggle('hidden', selectedValue !== 'custom-script');
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
        return Array.from(document.querySelectorAll('.mapping-row')).map(row => ({
            if: row.querySelector('.if-input').value,
            operator: row.querySelector('.operator-select').value,
            then: row.querySelector('.then-input').value,
        })).filter(m => m.if);
    }

    _scanScriptForInputs() {
        const script = document.getElementById('custom-script-editor').value;
        const regex = /\b(?:let|const|var)\s+([a-zA-Z_$][a-zA-Z0-9_$]*)(?!\s*=)/g;
        const newInputs = new Set();
        let match;
        while ((match = regex.exec(script)) !== null) {
            if (match[1] !== 'getWorksheet' && match[1] !== 'sourceColumnValue') {
                newInputs.add(match[1]);
            }
        }
        const updatedScriptInputs = {};
        newInputs.forEach(name => { updatedScriptInputs[name] = this.currentScriptInputs[name] || ''; });
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
            this.currentScriptInputs[input.dataset.varname] = input.value.trim();
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
        event.target.value = '';
    }

    async showSendConfirmModal() {
        if (this.appContext.studentDataCache.length === 0) {
            document.getElementById('status').textContent = 'Please select recipients before sending.';
            document.getElementById('status').style.color = 'orange';
            return;
        }
        const count = this.appContext.studentDataCache.length;
        document.getElementById('send-confirm-message').textContent = `You are about to send ${count} ${count === 1 ? 'email' : 'emails'}. Do you want to proceed?`;
        this.show('send-confirm-modal');
    }
}

