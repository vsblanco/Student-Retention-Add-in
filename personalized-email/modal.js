// V-1.1 - 2025-09-17 - 11:57 AM EDT

// A mapping of modal IDs to their corresponding wrapper elements.
const MODAL_ELEMENTS = {
    'example': 'example-modal',
    'payload': 'payload-modal',
    'templates': 'templates-modal',
    'save-template': 'save-template-modal',
    'custom-param': 'custom-param-modal',
    'manage-custom-params': 'manage-custom-params-modal',
    'send-confirm': 'send-confirm-modal'
};

/**
 * Manages the state and interactions of all modals in the task pane.
 */
export default class ModalManager {
    constructor(appContext) {
        this.appContext = appContext;
        this.activeModals = [];
        this.currentCustomParam = null;
        this.miniQuillEditor = null; // To hold the Quill instance for template editing.

        this._bindUIElements();
        this._bindEventListeners();
    }

    /**
     * Caches references to frequently used DOM elements.
     */
    _bindUIElements() {
        this.elements = {
            // Main UI buttons
            templatesButton: document.getElementById('templates-button'),
            showExampleButton: document.getElementById('show-example-button'),
            showPayloadButton: document.getElementById('show-payload-button'),
            createCustomParamButton: document.getElementById('create-custom-param-button'),

            // Custom Param Modal elements
            customParamModalTitle: document.getElementById('custom-param-modal-title'),
            manageCustomParamsButton: document.getElementById('manage-custom-params-button'),
            paramNameInput: document.getElementById('param-name'),
            paramSourceColumnInput: document.getElementById('param-source-column'),
            paramMappingContainer: document.getElementById('param-mapping-container'),
            addMappingButton: document.getElementById('add-mapping-button'),
            saveParamStatus: document.getElementById('save-param-status'),

            // New elements for Office Script parameters
            paramTypeMappingRadio: document.getElementById('param-type-mapping'),
            paramTypeScriptRadio: document.getElementById('param-type-script'),
            mappingSection: document.getElementById('param-logic-mapping-section'),
            scriptSection: document.getElementById('param-logic-script-section'),
            sourceColumnDesc: document.getElementById('param-source-column-desc'),
            officeScriptSelect: document.getElementById('param-office-script')
        };
    }

    /**
     * Attaches event listeners to UI elements.
     */
    _bindEventListeners() {
        // Modal Triggers
        this.elements.templatesButton.onclick = () => this.showTemplatesModal();
        this.elements.showExampleButton.onclick = () => this.showExampleModal();
        this.elements.showPayloadButton.onclick = () => this.showPayloadModal();
        this.elements.createCustomParamButton.onclick = () => this.showCustomParamModal(null);
        this.elements.manageCustomParamsButton.onclick = () => this.showManageCustomParamsModal();
        
        // Custom Param Modal listeners
        document.getElementById('save-custom-param-button').onclick = () => this._saveCustomParameter();
        this.elements.addMappingButton.onclick = () => this._addMappingRow();

        // New listeners for parameter type selection
        this.elements.paramTypeMappingRadio.addEventListener('change', this._toggleParamLogicType.bind(this));
        this.elements.paramTypeScriptRadio.addEventListener('change', this._toggleParamLogicType.bind(this));


        // General Modal Close Buttons
        this._bindCloseAction('close-example-modal-button', 'example');
        this._bindCloseAction('close-payload-modal-button', 'payload');
        this._bindCloseAction('close-templates-modal-button', 'templates');
        this._bindCloseAction('cancel-save-template-button', 'save-template');
        this._bindCloseAction('cancel-custom-param-button', 'custom-param');
        this._bindCloseAction('close-manage-custom-params-button', 'manage-custom-params');
        this._bindCloseAction('cancel-send-button', 'send-confirm');
        
        // Other Listeners
        document.getElementById('confirm-send-button').onclick = () => this.appContext.executeSend();
    }
    
    // --- Public Methods ---

    /**
     * Displays a modal by its key and pushes it to the active stack.
     * @param {string} modalKey - The key of the modal to show (e.g., 'example').
     */
    show(modalKey) {
        const modalId = MODAL_ELEMENTS[modalKey];
        if (modalId) {
            document.getElementById(modalId).classList.remove('hidden');
            this.activeModals.push(modalKey);
        }
    }

    /**
     * Hides a modal by its key and removes it from the active stack.
     * @param {string} modalKey - The key of the modal to hide.
     */
    hide(modalKey) {
        const modalId = MODAL_ELEMENTS[modalKey];
        if (modalId) {
            document.getElementById(modalId).classList.add('hidden');
            this.activeModals = this.activeModals.filter(m => m !== modalKey);
        }
    }

    /**
     * Hides the topmost active modal.
     */
    hideTop() {
        if (this.activeModals.length > 0) {
            const topModalKey = this.activeModals.pop();
            this.hide(topModalKey);
        }
    }

    showSendConfirmModal = async () => {
        try {
            await this.appContext.getStudentData(); // Ensure student data is loaded
            const studentCount = this.appContext.getStudentData.length;
            document.getElementById('send-confirm-message').textContent = `You are about to send personalized emails to ${studentCount} student(s). Do you want to proceed?`;
            this.show('send-confirm');
        } catch (error) {
            // Error is already handled and displayed in the status bar by getStudentData
        }
    }

    showExampleModal = async () => {
        try {
            const studentData = await this.appContext.getStudentData();
            if (studentData.length === 0) {
                this.appContext.status.textContent = 'No student data available to generate an example.';
                this.appContext.status.style.color = 'orange';
                return;
            }
            const firstStudent = studentData[0];
            const fromTemplate = document.getElementById('email-from').value;
            const subjectTemplate = document.getElementById('email-subject').value;
            const bodyTemplate = this.appContext.quill.root.innerHTML;

            document.getElementById('example-from').textContent = this.appContext.renderTemplate(fromTemplate, firstStudent);
            document.getElementById('example-to').textContent = firstStudent.StudentEmail || '(No Email Found)';
            document.getElementById('example-subject').textContent = this.appContext.renderTemplate(subjectTemplate, firstStudent);
            document.getElementById('example-cc').textContent = this.appContext.renderCCTemplate(this.appContext.ccRecipients, firstStudent);
            document.getElementById('example-body').innerHTML = this.appContext.renderTemplate(bodyTemplate, firstStudent);

            this.show('example');
        } catch (error) {
            console.error("Could not generate example:", error);
        }
    }
    
    showPayloadModal = async () => {
        // Implementation for showing payload modal
    }

    showTemplatesModal = async () => {
        // Implementation for showing templates modal
    }
    
    showCustomParamModal = (paramToEdit = null) => {
        this.currentCustomParam = paramToEdit;
        this._clearCustomParamForm();

        // Reset the UI to the default state (Value Mapping selected)
        this.elements.paramTypeMappingRadio.checked = true;
        this._toggleParamLogicType();

        // Fetch the available Office Scripts and populate the dropdown
        this._fetchAndPopulateOfficeScripts();
        
        if (paramToEdit) {
            this.elements.customParamModalTitle.textContent = "Edit Custom Parameter";
            this._populateCustomParamForm(paramToEdit);
        } else {
            this.elements.customParamModalTitle.textContent = "Create Custom Parameter";
        }
        
        this.show('custom-param');
    }

    showManageCustomParamsModal = async () => {
        // Implementation for managing custom parameters
    }

    // --- Private Helper Methods ---

    /**
     * Binds a click event to an element to close a specific modal.
     * @param {string} elementId - The ID of the clickable element.
     * @param {string} modalKey - The key of the modal to close.
     */
    _bindCloseAction(elementId, modalKey) {
        const element = document.getElementById(elementId);
        if (element) {
            element.onclick = () => this.hide(modalKey);
        }
    }
    
    _clearCustomParamForm() {
        this.elements.paramNameInput.value = '';
        this.elements.paramSourceColumnInput.value = '';
        this.elements.paramMappingContainer.innerHTML = '';
        this.setSaveParamStatus('');
    }

    _populateCustomParamForm(param) {
        this.elements.paramNameInput.value = param.name;
        this.elements.paramSourceColumnInput.value = param.sourceColumn;
        
        if (param.type === 'script') {
            this.elements.paramTypeScriptRadio.checked = true;
            // We need to wait for scripts to be fetched before setting the value
            this._fetchAndPopulateOfficeScripts().then(() => {
                this.elements.officeScriptSelect.value = param.scriptId;
            });
        } else { // Default to 'mapping' for older params or explicitly set
            this.elements.paramTypeMappingRadio.checked = true;
            if (param.mappings) {
                param.mappings.forEach(m => this._addMappingRow(m));
            }
        }
        this._toggleParamLogicType(); // Update UI visibility
    }

    /**
     * Toggles the visibility of parameter logic sections based on radio button selection.
     */
    _toggleParamLogicType() {
        if (this.elements.paramTypeScriptRadio.checked) {
            this.elements.mappingSection.classList.add('hidden');
            this.elements.scriptSection.classList.remove('hidden');
            this.elements.sourceColumnDesc.textContent = "This column's value for a student will be passed as an argument to the selected script.";
        } else { // Mapping is checked
            this.elements.mappingSection.classList.remove('hidden');
            this.elements.scriptSection.classList.add('hidden');
            this.elements.sourceColumnDesc.textContent = "The column in your sheet that this parameter will read from.";
        }
    }

    /**
     * Fetches Office Scripts from the workbook and populates the dropdown.
     */
    async _fetchAndPopulateOfficeScripts() {
        this.elements.officeScriptSelect.innerHTML = '<option value="">Loading scripts...</option>';
        this.elements.officeScriptSelect.disabled = true;
        try {
            await Excel.run(async (context) => {
                const scripts = context.workbook.scripts.load("items");
                await context.sync();
                
                this.elements.officeScriptSelect.innerHTML = '<option value="">-- Select a script --</option>';
                if (scripts.items.length === 0) {
                    this.elements.officeScriptSelect.innerHTML = '<option value="">No scripts found in workbook</option>';
                } else {
                    scripts.items.forEach(script => {
                        const option = document.createElement('option');
                        option.value = script.id;
                        option.textContent = script.name;
                        this.elements.officeScriptSelect.appendChild(option);
                    });
                    this.elements.officeScriptSelect.disabled = false;
                }
            });
        } catch (error) {
            console.error("Error fetching Office Scripts:", error);
            this.elements.officeScriptSelect.innerHTML = '<option value="">Error loading scripts</option>';
        }
    }
    
    _addMappingRow(mapping = { if: '', operator: 'eq', then: '' }) {
       // Implementation for adding a mapping row
    }

    _getMappingsFromDOM() {
        // Implementation to read mapping data from the UI
        return [];
    }

    async _saveCustomParameter() {
        const paramName = this.elements.paramNameInput.value.trim().replace(/\s+/g, '');
        const sourceColumn = this.elements.paramSourceColumnInput.value.trim();

        if (!paramName) {
            this.setSaveParamStatus('Parameter Name is required.', true);
            return;
        }

        const newParam = {
            id: this.currentCustomParam ? this.currentCustomParam.id : `param_${Date.now()}`,
            name: paramName,
            sourceColumn: sourceColumn
        };
        
        if (this.elements.paramTypeScriptRadio.checked) {
            const scriptSelect = this.elements.officeScriptSelect;
            if (!scriptSelect.value) {
                this.setSaveParamStatus('Please select a script.', true);
                return;
            }
            newParam.type = 'script';
            newParam.scriptId = scriptSelect.value;
            newParam.scriptName = scriptSelect.options[scriptSelect.selectedIndex].text;
            newParam.mappings = []; // Ensure mappings array is empty for script type
        } else {
            newParam.type = 'mapping';
            newParam.mappings = this._getMappingsFromDOM();
        }

        this.setSaveParamStatus('Saving...', false);

        try {
            let params = await this.appContext.getCustomParameters();
            if (this.currentCustomParam) {
                const index = params.findIndex(p => p.id === this.currentCustomParam.id);
                if (index !== -1) params[index] = newParam;
            } else {
                params.push(newParam);
            }
            await this.appContext.saveCustomParameters(params);
            await this.appContext.loadCustomParameters();
            this.appContext.populateParameterButtons();
            
            this.setSaveParamStatus('Saved successfully!', false);
            setTimeout(() => {
                this.hide('custom-param');
            }, 1000);
        } catch (error) {
            this.setSaveParamStatus('Error saving parameter.', true);
            console.error("Failed to save custom parameter:", error);
        }
    }

    setSaveParamStatus(message, isError = false) {
        this.elements.saveParamStatus.textContent = message;
        this.elements.saveParamStatus.style.color = isError ? 'red' : 'green';
    }
}
