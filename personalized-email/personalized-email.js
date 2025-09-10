import { DOM_IDS, SETTINGS_KEYS } from './constants.js';
import * as ui from './ui.js';
import * as data from './data.js';
import { updateState, getState } from './state.js';
import { Modal } from './modal.js';

// --- Modal Instances ---
let exampleModal, payloadModal, sendConfirmModal, templatesModal, saveTemplateModal, customParamModal, manageCustomParamsModal, randomizeModal;

// --- Initialization ---

/**
 * This is the main entry point for the application.
 * It's wrapped in a DOM readiness check to ensure all HTML elements are available.
 */
function startApp() {
    console.log("[LOG] DOM is ready. Initializing application.");
    ui.initializeQuill();
    initializeModals();
    setupEventListeners();

    // Check for an existing connection and set the initial view.
    data.checkConnection().then(connection => {
        if (connection) {
            updateState('powerAutomateUrl', connection.url);
            ui.showView('composer');
            // Load custom parameters and populate the UI
            data.getSettings(SETTINGS_KEYS.CUSTOM_PARAMS).then(customParams => {
                updateState('customParameters', customParams);
                ui.populateParameterButtons();
            });
        } else {
            ui.showView('setup');
        }
    }).catch(error => {
        console.error("Initialization check failed:", error);
        ui.updateStatus("Could not check connection.", "red");
    });
}

// FIX: Implement a robust, two-stage initialization.
// Stage 1: Wait for the Office host to be ready.
Office.onReady((info) => {
    if (info.host === Office.HostType.Excel) {
        console.log("[LOG] Office is ready. Waiting for DOM.");
        // Stage 2: Wait for the DOM to be fully loaded and parsed.
        if (document.readyState === 'loading') {
            document.addEventListener('DOMContentLoaded', startApp);
        } else {
            startApp();
        }
    }
});

function initializeModals() {
    exampleModal = new Modal(DOM_IDS.EXAMPLE_MODAL, { closeButtonId: DOM_IDS.CLOSE_EXAMPLE_MODAL_BUTTON });
    payloadModal = new Modal(DOM_IDS.PAYLOAD_MODAL, { closeButtonId: DOM_IDS.CLOSE_PAYLOAD_MODAL_BUTTON });
    sendConfirmModal = new Modal(DOM_IDS.SEND_CONFIRM_MODAL, { closeButtonId: DOM_IDS.CANCEL_SEND_BUTTON });
    templatesModal = new Modal(DOM_IDS.TEMPLATES_MODAL, { closeButtonId: DOM_IDS.CLOSE_TEMPLATES_MODAL_BUTTON });
    saveTemplateModal = new Modal(DOM_IDS.SAVE_TEMPLATE_MODAL, { closeButtonId: DOM_IDS.CANCEL_SAVE_TEMPLATE_BUTTON });
    customParamModal = new Modal(DOM_IDS.CUSTOM_PARAM_MODAL, { closeButtonId: DOM_IDS.CANCEL_CUSTOM_PARAM_BUTTON });
    manageCustomParamsModal = new Modal(DOM_IDS.MANAGE_CUSTOM_PARAMS_MODAL, { closeButtonId: DOM_IDS.CLOSE_MANAGE_PARAMS_BUTTON });
    randomizeModal = new Modal(DOM_IDS.RANDOMIZE_MODAL, { closeButtonId: DOM_IDS.CANCEL_RANDOMIZE_BUTTON });
}


// --- Event Listeners Setup ---

function setupEventListeners() {
    const listeners = {
        // Main Actions
        [DOM_IDS.CREATE_CONNECTION_BUTTON]: handleCreateConnection,
        [DOM_IDS.SEND_EMAIL_BUTTON]: handleSendEmail,
        [DOM_IDS.SHOW_EXAMPLE_BUTTON]: handleShowExample,
        [DOM_IDS.SHOW_PAYLOAD_BUTTON]: handleShowPayload,

        // Modal Triggers
        [DOM_IDS.TEMPLATES_BUTTON]: handleShowTemplates,
        [DOM_IDS.CREATE_CUSTOM_PARAM_BUTTON]: () => ui.showCustomParamModal(),
        [DOM_IDS.MANAGE_CUSTOM_PARAMS_BUTTON]: handleManageCustomParams,
        [DOM_IDS.RANDOMIZE_PARAMETER_BUTTON]: handleShowRandomizeModal,

        // Modal Actions
        [DOM_IDS.TOGGLE_PAYLOAD_SCHEMA_BUTTON]: ui.togglePayloadSchemaView,
        [DOM_IDS.CONFIRM_SEND_BUTTON]: handleExecuteSend,
        [DOM_IDS.SAVE_CURRENT_TEMPLATE_BUTTON]: handleSaveTemplate,
        [DOM_IDS.CONFIRM_SAVE_TEMPLATE_BUTTON]: handleConfirmSaveTemplate,
        [DOM_IDS.SAVE_CUSTOM_PARAM_BUTTON]: handleSaveCustomParameter,
        [DOM_IDS.ADD_MAPPING_BUTTON]: () => ui.addMappingRow(),
        [DOM_IDS.ADD_RANDOMIZE_OPTION_BUTTON]: () => ui.addRandomizeOption(),
        [DOM_IDS.CONFIRM_RANDOMIZE_BUTTON]: handleConfirmRandomize,


        // Other UI
        [DOM_IDS.RECIPIENT_LIST]: ui.toggleCustomSheetInput
    };

    for (const id in listeners) {
        const element = document.getElementById(id);
        if (element) {
            element.onclick = listeners[id];
        } else {
            console.warn(`Element with ID "${id}" not found. Cannot attach event listener.`);
        }
    }

    ui.setupPillboxInputs();

    // Close dropdowns when clicking outside
    document.addEventListener('click', (e) => {
        if (!e.target.closest('.randomize-tag-wrapper') && !e.target.closest('.condition-tag-wrapper')) {
           document.querySelectorAll('.randomize-panel, .condition-panel').forEach(p => p.style.display = 'none');
           document.querySelectorAll('.randomize-arrow, .condition-arrow').forEach(a => a.classList.remove('open'));
       }
   });
}


// --- Event Handlers ---

function handleCreateConnection() {
    const url = ui.getPowerAutomateUrl();
    if (!url.startsWith('http')) {
        ui.updateStatus("Please enter a valid HTTP URL.", 'red', true);
        return;
    }

    ui.updateStatus("Creating connection...", 'gray', true);
    data.createConnection(url).then(newConnection => {
        updateState('powerAutomateUrl', newConnection.url);
        ui.updateStatus("Connection created successfully!", 'green', true);
        setTimeout(() => ui.showView('composer'), 1500);
    }).catch(error => {
        ui.updateStatus(`Error: ${error.message}`, 'red', true);
    });
}

function handleShowPayload() {
    console.log("[LOG] 'Show Payload' button clicked.");
    ui.updateStatus("Generating payload preview...");

    const template = ui.getEmailTemplateFromDOM();
    if (!template.from) {
        ui.updateStatus("Please enter a 'From' address before generating the payload.", 'orange');
        return;
    }

    const sheetName = ui.getSelectedSheetName();
    const { customParameters } = getState();

    data.getStudentData(sheetName, customParameters).then(students => {
        updateState('students', students);

        if (students.length === 0) {
            ui.updateStatus('No students found to generate a payload.', 'orange');
            ui.populatePayloadModal([]);
            payloadModal.show();
            return;
        }

        const payload = data.buildPayload(students, template);
        ui.populatePayloadModal(payload);
        payloadModal.show();
        ui.updateStatus("");
    }).catch(error => {
        ui.updateStatus(`Error: ${error.message}`, 'red');
        console.error("Error showing payload:", error);
    });
}

function handleShowExample() {
    console.log("[LOG] 'Example' button clicked.");
    ui.updateStatus("Generating example...");
    
    const template = ui.getEmailTemplateFromDOM();
    if (!template.from) {
        ui.updateStatus("Please enter a 'From' address to generate an example.", 'orange');
        return;
    }

    const sheetName = ui.getSelectedSheetName();
    const { customParameters } = getState();

    data.getStudentData(sheetName, customParameters).then(students => {
        updateState('students', students);

        if (students.length === 0) {
            ui.updateStatus('No students found to generate an example.', 'orange');
            return;
        }
        
        const validStudents = students.filter(s => s.StudentEmail);
        if (validStudents.length === 0) {
            ui.updateStatus('No students with a valid email address were found in the list.', 'orange');
            return;
        }
        
        const randomStudent = validStudents[Math.floor(Math.random() * validStudents.length)];
        const examplePayload = data.buildPayload([randomStudent], template);

        if (examplePayload.length === 0) {
             ui.updateStatus('Could not generate a valid example for the selected student. Try again.', 'orange');
             return;
        }
        
        ui.populateExampleModal(examplePayload[0]);
        exampleModal.show();
        ui.updateStatus("");
    }).catch(error => {
        ui.updateStatus(`Error: ${error.message}`, 'red');
        console.error("Error showing example:", error);
    });
}

function handleSendEmail() {
    console.log("[LOG] 'Send Email' button clicked.");
    ui.updateStatus("Preparing to send...");

    const template = ui.getEmailTemplateFromDOM();
    if (!template.from) {
        ui.updateStatus("Please enter a 'From' address before sending.", 'orange');
        return;
    }

    const sheetName = ui.getSelectedSheetName();
    const { customParameters } = getState();

    data.getStudentData(sheetName, customParameters).then(students => {
        updateState('students', students);

        if (students.length === 0) {
            ui.updateStatus('No students to send emails to.', 'orange');
            return;
        }
        
        const payload = data.buildPayload(students, template);
        if (payload.length === 0) {
            ui.updateStatus('No students with valid email addresses were found to send to.', 'orange');
            return;
        }
        
        ui.populateSendConfirmationModal(payload.length);
        sendConfirmModal.show();
    }).catch(error => {
        ui.updateStatus(`Error: ${error.message}`, 'red');
    });
}

function handleExecuteSend() {
    sendConfirmModal.hide();
    const { students, powerAutomateUrl } = getState();
    ui.updateStatus(`Sending ${students.length} emails...`);

    try {
        const template = ui.getEmailTemplateFromDOM();
        const payload = data.buildPayload(students, template);

        if (payload.length === 0) {
            ui.updateStatus('No valid students with "To" and "From" addresses found.', 'orange');
            return;
        }

        data.sendToPowerAutomate(powerAutomateUrl, payload).then(() => {
            ui.updateStatus(`Successfully sent ${payload.length} emails!`, 'green');
        }).catch(error => {
            ui.updateStatus(`Failed to send emails: ${error.message}`, 'red');
        });
    } catch (error) {
        ui.updateStatus(`Failed to send emails: ${error.message}`, 'red');
    }
}

// --- Template Handlers ---

function handleShowTemplates() {
    templatesModal.show();
    const handlers = {
        onLoad: handleLoadTemplate,
        onDelete: handleDeleteTemplate
    };
    data.getSettings(SETTINGS_KEYS.EMAIL_TEMPLATES).then(templates => {
        ui.populateTemplatesModal(templates, handlers);
    });
}

function handleSaveTemplate() {
    templatesModal.hide();
    ui.clearSaveTemplateForm();
    saveTemplateModal.show();
}

function handleConfirmSaveTemplate() {
    const formData = ui.getTemplateSaveForm();
    if (!formData.name || !formData.author) {
        ui.updateSaveTemplateStatus('Name and Author are required.', 'red');
        return;
    }
    ui.updateSaveTemplateStatus('Saving...');

    const templateContent = ui.getEmailTemplateFromDOM();
    const newTemplate = {
        id: 'template_' + new Date().getTime(),
        name: formData.name,
        author: formData.author,
        timestamp: new Date().toISOString(),
        ...templateContent
    };

    data.getSettings(SETTINGS_KEYS.EMAIL_TEMPLATES).then(templates => {
        templates.push(newTemplate);
        return data.saveSettings(SETTINGS_KEYS.EMAIL_TEMPLATES, templates);
    }).then(() => {
        ui.updateSaveTemplateStatus('Template saved!', 'green');
        setTimeout(() => saveTemplateModal.hide(), 1500);
    });
}

function handleLoadTemplate(templateId) {
    data.getSettings(SETTINGS_KEYS.EMAIL_TEMPLATES).then(templates => {
        const template = templates.find(t => t.id === templateId);
        if (template) {
            ui.loadTemplateIntoForm(template);
            templatesModal.hide();
        }
    });
}

function handleDeleteTemplate(templateId) {
    data.getSettings(SETTINGS_KEYS.EMAIL_TEMPLATES).then(templates => {
        const updatedTemplates = templates.filter(t => t.id !== templateId);
        return data.saveSettings(SETTINGS_KEYS.EMAIL_TEMPLATES, updatedTemplates);
    }).then(() => {
        handleShowTemplates(); // Refresh the modal content
    });
}

// --- Custom Parameter Handlers ---

function handleManageCustomParams() {
    customParamModal.hide();
    manageCustomParamsModal.show();
    const handlers = {
        onEdit: handleEditCustomParameter,
        onDelete: handleDeleteCustomParameter,
        onDuplicate: handleDuplicateCustomParameter
    };
    data.getSettings(SETTINGS_KEYS.CUSTOM_PARAMS).then(params => {
        ui.populateManageCustomParamsModal(params, handlers);
    });
}

function handleSaveCustomParameter() {
    const formData = ui.getCustomParamForm();
    if (!/^[a-zA-Z0-9]+$/.test(formData.name)) {
        ui.updateCustomParamStatus('Name must be alphanumeric with no spaces.', 'red');
        return;
    }
    
    ui.updateCustomParamStatus('Saving...');

    data.getSettings(SETTINGS_KEYS.CUSTOM_PARAMS).then(currentParams => {
        const existingParam = currentParams.find(p => p.name.toLowerCase() === formData.name.toLowerCase());
        if (existingParam && existingParam.id !== formData.id) {
            ui.updateCustomParamStatus('This parameter name is already in use.', 'red');
            return Promise.reject('Parameter name in use'); 
        }

        if (formData.id) { // Editing existing
            const index = currentParams.findIndex(p => p.id === formData.id);
            if (index > -1) currentParams[index] = { ...currentParams[index], ...formData };
        } else { // Creating new
            formData.id = 'sparam_' + new Date().getTime();
            currentParams.push(formData);
        }
        
        return data.saveSettings(SETTINGS_KEYS.CUSTOM_PARAMS, currentParams);
    }).then((savedParams) => {
        updateState('customParameters', savedParams);
        ui.populateParameterButtons();
        ui.updateCustomParamStatus('Parameter saved successfully!', 'green');
        setTimeout(() => customParamModal.hide(), 1500);
    }).catch(error => {
        if (error !== 'Parameter name in use') {
            ui.updateCustomParamStatus(`Error: ${error.message}`, 'red');
        }
    });
}

function handleEditCustomParameter(paramId) {
    data.getSettings(SETTINGS_KEYS.CUSTOM_PARAMS).then(params => {
        const param = params.find(p => p.id === paramId);
        if (param) {
            manageCustomParamsModal.hide();
            ui.showCustomParamModal(param);
        }
    });
}

function handleDeleteCustomParameter(paramId) {
    data.getSettings(SETTINGS_KEYS.CUSTOM_PARAMS).then(params => {
        const updatedParams = params.filter(p => p.id !== paramId);
        return data.saveSettings(SETTINGS_KEYS.CUSTOM_PARAMS, updatedParams);
    }).then((savedParams) => {
        updateState('customParameters', savedParams);
        ui.populateParameterButtons();
        handleManageCustomParams(); // Refresh list
    });
}

function handleDuplicateCustomParameter(paramId) {
    data.getSettings(SETTINGS_KEYS.CUSTOM_PARAMS).then(params => {
        const paramToDuplicate = params.find(p => p.id === paramId);
        if (!paramToDuplicate) return;

        const newParam = JSON.parse(JSON.stringify(paramToDuplicate));
        newParam.name = `${newParam.name}Copy`;
        newParam.id = 'sparam_' + new Date().getTime();
        params.push(newParam);

        return data.saveSettings(SETTINGS_KEYS.CUSTOM_PARAMS, params);
    }).then((savedParams) => {
        updateState('customParameters', savedParams);
        ui.populateParameterButtons();
        handleManageCustomParams(); // Refresh list
    });
}


// --- Special Parameter Handlers ---

function handleShowRandomizeModal() {
    ui.clearRandomizeModal();
    randomizeModal.show();
}

function handleConfirmRandomize() {
    const options = ui.getRandomizeOptions();
    if (options.length < 2) {
        // You might want to show an error to the user here
        console.warn("At least two options are required for randomization.");
        return;
    }
    const tag = `{Randomize|${options.join('|')}}`;
    ui.insertTextInQuill(tag);
    randomizeModal.hide();
}
