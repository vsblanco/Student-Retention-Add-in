// V-1.4 - 2025-09-11 - 11:49 AM EDT

/**
 * @fileoverview Manages all modal dialog interactions for the Personalized Email add-in.
 */
import { PAYLOAD_SCHEMA } from './constants.js';

// This class encapsulates the logic for showing, hiding, and handling actions within modals.
export default class ModalManager {
    /**
     * @param {object} app - The main application context.
     * @param {Quill} app.quill - The Quill editor instance.
     * @param {Function} app.getStudentData - Function to fetch student data.
     * @param {Function} app.renderTemplate - Function to render a template string.
     * @param {Function} app.renderCCTemplate - Function to render the CC template string.
     * @param {Function} app.getTemplates - Function to get saved templates.
     * @param {Function} app.saveTemplates - Function to save templates.
     * @param {Function} app.loadCustomParameters - Function to load custom parameters.
     * @param {Function} app.getCustomParameters - Function to get custom parameters.
     * @param {Function} app.saveCustomParameters - Function to save custom parameters.
     * @param {Function} app.populateParameterButtons - Function to update the parameter buttons UI.
     * @param {Function} app.executeSend - Function to execute the email sending process.
     * @param {Array} app.ccRecipients - A reference to the ccRecipients array.
     * @param {Function} app.renderCCPills - Function to re-render the CC pills UI.
     */
    constructor(app) {
        this.app = app;
        this.bindModalEventListeners();
    }

    bindModalEventListeners() {
        // Main Buttons that open modals
        document.getElementById('show-example-button').onclick = this.showExample.bind(this);
        document.getElementById('show-payload-button').onclick = this.showPayload.bind(this);
        document.getElementById('templates-button').onclick = this.showTemplatesModal.bind(this);
        document.getElementById('create-custom-param-button').onclick = this.showCustomParamModal.bind(this);

        // Modal Close Buttons
        document.getElementById('close-example-modal-button').onclick = () => document.getElementById('example-modal').classList.add('hidden');
        document.getElementById('close-payload-modal-button').onclick = () => document.getElementById('payload-modal').classList.add('hidden');
        document.getElementById('close-templates-modal-button').onclick = () => document.getElementById('templates-modal').classList.add('hidden');
        document.getElementById('cancel-save-template-button').onclick = () => document.getElementById('save-template-modal').classList.add('hidden');
        document.getElementById('cancel-send-button').onclick = () => document.getElementById('send-confirm-modal').classList.add('hidden');
        document.getElementById('cancel-custom-param-button').onclick = () => document.getElementById('custom-param-modal').classList.add('hidden');
        document.getElementById('close-manage-custom-params-button').onclick = () => document.getElementById('manage-custom-params-modal').classList.add('hidden');
        
        // Modal Action Buttons
        document.getElementById('toggle-payload-schema-button').onclick = this.togglePayloadView.bind(this);
        document.getElementById('save-current-template-button').onclick = this.showSaveTemplateModal.bind(this);
        document.getElementById('confirm-save-template-button').onclick = this.saveTemplate.bind(this);
        document.getElementById('confirm-send-button').onclick = this.app.executeSend;
        document.getElementById('save-custom-param-button').onclick = this.saveCustomParameter.bind(this);
        document.getElementById('add-mapping-button').onclick = this.addMappingRow.bind(this);
        document.getElementById('manage-custom-params-button').onclick = this.showManageCustomParamsModal.bind(this);
    }

    async showExample() {
        const status = document.getElementById('status');
        try {
            const studentData = await this.app.getStudentData();
    
            if (studentData.length === 0) {
                status.textContent = 'No students found to generate an example.';
                status.style.color = 'orange';
                return;
            }
    
            const randomStudent = studentData[Math.floor(Math.random() * studentData.length)];
            
            const fromTemplate = document.getElementById('email-from').value;
            const subjectTemplate = document.getElementById('email-subject').value;
            const bodyTemplate = this.app.quill.root.innerHTML;
    
            document.getElementById('example-from').textContent = this.app.renderTemplate(fromTemplate, randomStudent) || '[Not Specified]';
            document.getElementById('example-to').textContent = randomStudent.StudentEmail || '[No Email Found]';
            document.getElementById('example-cc').textContent = this.app.renderCCTemplate(this.app.ccRecipients, randomStudent) || '[Not Specified]';
            document.getElementById('example-subject').textContent = this.app.renderTemplate(subjectTemplate, randomStudent);
            document.getElementById('example-body').innerHTML = this.app.renderTemplate(bodyTemplate, randomStudent);
    
            document.getElementById('example-modal').classList.remove('hidden');
    
        } catch (error) {
            // Error message is already set by getStudentData
        }
    }
    
    async showPayload() {
        const status = document.getElementById('status');
        try {
            const studentData = await this.app.getStudentData();
    
            if (studentData.length === 0) {
                status.textContent = 'No students found to generate a payload.';
                status.style.color = 'orange';
                return;
            }
    
            const fromTemplate = document.getElementById('email-from').value;
            const subjectTemplate = document.getElementById('email-subject').value;
            const bodyTemplate = this.app.quill.root.innerHTML;
    
            const payload = studentData.map(student => ({
                from: this.app.renderTemplate(fromTemplate, student),
                to: student.StudentEmail || '',
                cc: this.app.renderCCTemplate(this.app.ccRecipients, student),
                subject: this.app.renderTemplate(subjectTemplate, student),
                body: this.app.renderTemplate(bodyTemplate, student)
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
    
    togglePayloadView() {
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

    async showSendConfirmModal() {
        const status = document.getElementById('status');
        try {
            const studentData = await this.app.getStudentData();
    
            if (studentData.length === 0) {
                status.textContent = 'No students to send emails to.';
                status.style.color = 'orange';
                return;
            }
            
            const confirmMessage = document.getElementById('send-confirm-message');
            confirmMessage.textContent = `You are about to send emails to ${studentData.length} student(s). Do you want to proceed?`;
            document.getElementById('send-confirm-modal').classList.remove('hidden');
    
        } catch (error) {
            // Error message is already set by getStudentData
        }
    }

    async showTemplatesModal() {
        const container = document.getElementById('templates-list-container');
        container.innerHTML = '<p class="text-gray-500">Loading templates...</p>';
        document.getElementById('templates-modal').classList.remove('hidden');
    
        const templates = await this.app.getTemplates();
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
            btn.onclick = () => this.loadTemplate(btn.dataset.id);
        });
        container.querySelectorAll('.delete-template-btn').forEach(btn => {
            btn.onclick = () => this.deleteTemplate(btn.dataset.id);
        });
    }
    
    showSaveTemplateModal() {
        document.getElementById('templates-modal').classList.add('hidden');
        document.getElementById('template-name').value = '';
        document.getElementById('template-author').value = ''; 
        document.getElementById('save-template-status').textContent = '';
        document.getElementById('save-template-modal').classList.remove('hidden');
    }
    
    async saveTemplate() {
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
            cc: this.app.ccRecipients,
            body: this.app.quill.root.innerHTML
        };
    
        const templates = await this.app.getTemplates();
        templates.push(newTemplate);
        await this.app.saveTemplates(templates);
    
        status.textContent = 'Template saved!';
        status.style.color = 'green';
        setTimeout(() => {
            document.getElementById('save-template-modal').classList.add('hidden');
        }, 1500);
    }
    
    async loadTemplate(templateId) {
        const templates = await this.app.getTemplates();
        const template = templates.find(t => t.id === templateId);
        if (template) {
            document.getElementById('email-from').value = template.from || '';
            document.getElementById('email-subject').value = template.subject;
            this.app.ccRecipients.length = 0; // Clear the array
            Array.prototype.push.apply(this.app.ccRecipients, template.cc || []);
            this.app.renderCCPills();
            this.app.quill.root.innerHTML = template.body;
            document.getElementById('templates-modal').classList.add('hidden');
        }
    }
    
    async deleteTemplate(templateId) {
        let templates = await this.app.getTemplates();
        templates = templates.filter(t => t.id !== templateId);
        await this.app.saveTemplates(templates);
        await this.showTemplatesModal(); // Refresh the list
    }

    async showCustomParamModal() {
        document.getElementById('param-source-column').value = '';
        document.getElementById('param-name').value = '';
        document.getElementById('param-default-value').value = '';
        document.getElementById('param-mapping-container').innerHTML = '';
        document.getElementById('save-param-status').textContent = '';
        
        document.getElementById('custom-param-modal').classList.remove('hidden');
    }
    
    addMappingRow() {
        const container = document.getElementById('param-mapping-container');
        const div = document.createElement('div');
        div.className = 'flex items-center gap-2 mapping-row';
        div.innerHTML = `
            <span class="text-sm text-gray-500">If cell is</span>
            <input type="text" class="mapping-if flex-1 px-2 py-1 border border-gray-300 rounded-md text-sm" placeholder="e.g., Bob">
            <span class="text-sm text-gray-500">then value is</span>
            <input type="text" class="mapping-then flex-1 px-2 py-1 border border-gray-300 rounded-md text-sm" placeholder="e.g., bobjones@gmail.com">
            <button class="remove-mapping-btn text-red-500 hover:text-red-700 text-lg">&times;</button>
        `;
        div.querySelector('.remove-mapping-btn').onclick = () => div.remove();
        container.appendChild(div);
    }
    
    async saveCustomParameter() {
        const status = document.getElementById('save-param-status');
        const nameInput = document.getElementById('param-name');
        const name = nameInput.value.trim();
    
        // Validation
        if (!/^[a-zA-Z0-9]+$/.test(name)) {
            status.textContent = 'Name must be alphanumeric with no spaces.';
            status.style.color = 'red';
            return;
        }
        if (this.app.standardParameters.includes(name) || this.app.customParameters.find(p => p.name.toLowerCase() === name.toLowerCase())) {
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
    
        const newParam = {
            id: 'cparam_' + new Date().getTime(),
            name,
            sourceColumn,
            defaultValue,
            mappings
        };
    
        status.textContent = 'Saving...';
        status.style.color = 'gray';
    
        const currentParams = await this.app.getCustomParameters();
        currentParams.push(newParam);
        await this.app.saveCustomParameters(currentParams);
        
        await this.app.loadCustomParameters(); 
        await this.app.populateParameterButtons();
    
        status.textContent = 'Parameter saved successfully!';
        status.style.color = 'green';
        setTimeout(() => {
            document.getElementById('custom-param-modal').classList.add('hidden');
        }, 1500);
    }
    
    async showManageCustomParamsModal() {
        document.getElementById('custom-param-modal').classList.add('hidden');
        const listContainer = document.getElementById('manage-custom-params-list');
        listContainer.innerHTML = '<p class="text-gray-500">Loading...</p>';
        document.getElementById('manage-custom-params-modal').classList.remove('hidden');
    
        const params = await this.app.getCustomParameters();
        listContainer.innerHTML = '';
        if (params.length === 0) {
            listContainer.innerHTML = '<p class="text-gray-500 text-center">No custom parameters created yet.</p>';
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
                    <button data-id="${param.id}" class="delete-param-btn px-3 py-1 bg-red-100 text-red-800 text-xs font-semibold rounded-md hover:bg-red-200">Delete</button>
                </div>
                <div class="mt-2 text-sm">${mappingsHtml}</div>
            `;
            listContainer.appendChild(div);
        });
    
        listContainer.querySelectorAll('.delete-param-btn').forEach(btn => {
            btn.onclick = () => this.deleteCustomParameter(btn.dataset.id);
        });
    }
    
    async deleteCustomParameter(paramId) {
        let params = await this.app.getCustomParameters();
        params = params.filter(p => p.id !== paramId);
        await this.app.saveCustomParameters(params);
        await this.app.loadCustomParameters();
        await this.app.populateParameterButtons();
        await this.showManageCustomParamsModal();
    }
}
