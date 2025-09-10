import { DOM_IDS, SETTINGS_KEYS } from './constants.js';
import * as ui from './ui.js';
import * as data from './data.js';
import { updateState, getState } from './state.js';
import { Modal } from './modal.js';

// --- Modal Instances ---
let exampleModal, payloadModal, sendConfirmModal, templatesModal, saveTemplateModal, customParamModal, manageCustomParamsModal;

// --- Initialization ---

Office.onReady((info) => {
    if (info.host === Office.HostType.Excel) {
        initializeApplication();
    }
});

async function initializeApplication() {
    ui.initializeQuill();
    initializeModals();
    setupEventListeners();

    const connection = await data.checkConnection();
    if (connection) {
        updateState('powerAutomateUrl', connection.url);
        ui.showView('composer');
        const customParams = await data.getSettings(SETTINGS_KEYS.CUSTOM_PARAMS);
        updateState('customParameters', customParams);
        ui.populateParameterButtons();
    } else {
        ui.showView('setup');
    }
}

function initializeModals() {
    exampleModal = new Modal(DOM_IDS.EXAMPLE_MODAL, { closeButtonId: DOM_IDS.CLOSE_EXAMPLE_MODAL_BUTTON });
    payloadModal = new Modal(DOM_IDS.PAYLOAD_MODAL, { closeButtonId: DOM_IDS.CLOSE_PAYLOAD_MODAL_BUTTON });
    sendConfirmModal = new Modal(DOM_IDS.SEND_CONFIRM_MODAL, { closeButtonId: DOM_IDS.CANCEL_SEND_BUTTON });
    templatesModal = new Modal(DOM_IDS.TEMPLATES_MODAL, { closeButtonId: DOM_IDS.CLOSE_TEMPLATES_MODAL_BUTTON });
    saveTemplateModal = new Modal(DOM_IDS.SAVE_TEMPLATE_MODAL, { closeButtonId: DOM_IDS.CANCEL_SAVE_TEMPLATE_BUTTON });
    customParamModal = new Modal(DOM_IDS.CUSTOM_PARAM_MODAL, { closeButtonId: DOM_IDS.CANCEL_CUSTOM_PARAM_BUTTON });
    manageCustomParamsModal = new Modal(DOM_IDS.MANAGE_CUSTOM_PARAMS_MODAL, { closeButtonId: DOM_IDS.CLOSE_MANAGE_PARAMS_BUTTON });
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

        // Modal Actions
        [DOM_IDS.TOGGLE_PAYLOAD_SCHEMA_BUTTON]: ui.togglePayloadSchemaView,
        [DOM_IDS.CONFIRM_SEND_BUTTON]: handleExecuteSend,
        [DOM_IDS.SAVE_CURRENT_TEMPLATE_BUTTON]: handleSaveTemplate,
        [DOM_IDS.CONFIRM_SAVE_TEMPLATE_BUTTON]: handleConfirmSaveTemplate,
        [DOM_IDS.SAVE_CUSTOM_PARAM_BUTTON]: handleSaveCustomParameter,
        [DOM_IDS.ADD_MAPPING_BUTTON]: () => ui.addMappingRow(),

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

async function handleCreateConnection() {
    const url = ui.getPowerAutomateUrl();
    if (!url.startsWith('http')) {
        ui.updateStatus("Please enter a valid HTTP URL.", 'red', true);
        return;
    }

    ui.updateStatus("Creating connection...", 'gray', true);
    try {
        const newConnection = await data.createConnection(url);
        updateState('powerAutomateUrl', newConnection.url);
        ui.updateStatus("Connection created successfully!", 'green', true);
        setTimeout(() => ui.showView('composer'), 1500);
    } catch (error) {
        ui.updateStatus(`Error: ${error.message}`, 'red', true);
    }
}

async function handleShowPayload() {
    console.log("[LOG] 'Show Payload' button clicked.");
    ui.updateStatus("Generating payload preview...");
    try {
        // FIX: Add validation to ensure 'From' field is not empty.
        const template = ui.getEmailTemplateFromDOM();
        if (!template.from) {
            ui.updateStatus("Please enter a 'From' address before generating the payload.", 'orange');
            return;
        }

        const sheetName = ui.getSelectedSheetName();
        const { customParameters } = getState();
        const students = await data.getStudentData(sheetName, customParameters);
        updateState('students', students);

        if (students.length === 0) {
            ui.updateStatus('No students found to generate a payload.', 'orange');
            // Still show the modal with an empty payload for debugging.
            ui.populatePayloadModal([]);
            payloadModal.show();
            return;
        }

        const payload = data.buildPayload(students, template);

        ui.populatePayloadModal(payload);
        payloadModal.show();
        ui.updateStatus("");
    } catch (error) {
        ui.updateStatus(`Error: ${error.message}`, 'red');
        console.error("Error showing payload:", error);
    }
}

async function handleShowExample() {
    console.log("[LOG] 'Example' button clicked.");
    ui.updateStatus("Generating example...");
    try {
        // FIX: Add validation to ensure 'From' field is not empty.
        const template = ui.getEmailTemplateFromDOM();
        if (!template.from) {
            ui.updateStatus("Please enter a 'From' address to generate an example.", 'orange');
            return;
        }

        const sheetName = ui.getSelectedSheetName();
        const { customParameters } = getState();
        const students = await data.getStudentData(sheetName, customParameters);
        updateState('students', students);

        if (students.length === 0) {
            ui.updateStatus('No students found to generate an example.', 'orange');
            return;
        }
        
        // Find a random student that has a valid email address
        const validStudents = students.filter(s => s.StudentEmail);
        if (validStudents.length === 0) {
            ui.updateStatus('No students with a valid email address were found in the list.', 'orange');
            return;
        }
        
        const randomStudent = validStudents[Math.floor(Math.random() * validStudents.length)];
        
        // Build a single-item payload to get the rendered example
        const examplePayload = data.buildPayload([randomStudent], template);

        if (examplePayload.length === 0) {
             ui.updateStatus('Could not generate a valid example for the selected student. Try again.', 'orange');
             return;
        }
        
        ui.populateExampleModal(examplePayload[0]);
        exampleModal.show();
        ui.updateStatus("");
    } catch (error) {
        ui.updateStatus(`Error: ${error.message}`, 'red');
        console.error("Error showing example:", error);
    }
}

async function handleSendEmail() {
    console.log("[LOG] 'Send Email' button clicked.");
    ui.updateStatus("Preparing to send...");
    try {
        // FIX: Add validation to ensure 'From' field is not empty.
        const template = ui.getEmailTemplateFromDOM();
        if (!template.from) {
            ui.updateStatus("Please enter a 'From' address before sending.", 'orange');
            return;
        }

        const sheetName = ui.getSelectedSheetName();
        const { customParameters } = getState();
        const students = await data.getStudentData(sheetName, customParameters);
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
    } catch (error) {
        ui.updateStatus(`Error: ${error.message}`, 'red');
    }
}

async function handleExecuteSend() {
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

        await data.sendToPowerAutomate(powerAutomateUrl, payload);
        ui.updateStatus(`Successfully sent ${payload.length} emails!`, 'green');
    } catch (error) {
        ui.updateStatus(`Failed to send emails: ${error.message}`, 'red');
    }
}

// --- Template Handlers ---

async function handleShowTemplates() {
    templatesModal.show();
    const templates = await data.getSettings(SETTINGS_KEYS.EMAIL_TEMPLATES);
    const handlers = {
        onLoad: handleLoadTemplate,
        onDelete: handleDeleteTemplate
    };
    ui.populateTemplatesModal(templates, handlers);
}

function handleSaveTemplate() {
    templatesModal.hide();
    ui.clearSaveTemplateForm();
    saveTemplateModal.show();
}

async function handleConfirmSaveTemplate() {
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

    const templates = await data.getSettings(SETTINGS_KEYS.EMAIL_TEMPLATES);
    templates.push(newTemplate);
    await data.saveSettings(SETTINGS_KEYS.EMAIL_TEMPLATES, templates);

    ui.updateSaveTemplateStatus('Template saved!', 'green');
    setTimeout(() => saveTemplateModal.hide(), 1500);
}

async function handleLoadTemplate(templateId) {
    const templates = await data.getSettings(SETTINGS_KEYS.EMAIL_TEMPLATES);
    const template = templates.find(t => t.id === templateId);
    if (template) {
        ui.loadTemplateIntoForm(template);
        templatesModal.hide();
    }
}

async function handleDeleteTemplate(templateId) {
    let templates = await data.getSettings(SETTINGS_KEYS.EMAIL_TEMPLATES);
    templates = templates.filter(t => t.id !== templateId);
    await data.saveSettings(SETTINGS_KEYS.EMAIL_TEMPLATES, templates);
    await handleShowTemplates(); // Refresh the modal content
}

// --- Custom Parameter Handlers ---

async function handleManageCustomParams() {
    customParamModal.hide();
    manageCustomParamsModal.show();
    const params = await data.getSettings(SETTINGS_KEYS.CUSTOM_PARAMS);
    const handlers = {
        onEdit: handleEditCustomParameter,
        onDelete: handleDeleteCustomParameter,
        onDuplicate: handleDuplicateCustomParameter
    };
    ui.populateManageCustomParamsModal(params, handlers);
}

async function handleSaveCustomParameter() {
    const formData = ui.getCustomParamForm();
    if (!/^[a-zA-Z0-9]+$/.test(formData.name)) {
        ui.updateCustomParamStatus('Name must be alphanumeric with no spaces.', 'red');
        return;
    }
    
    let currentParams = await data.getSettings(SETTINGS_KEYS.CUSTOM_PARAMS);
    const existingParam = currentParams.find(p => p.name.toLowerCase() === formData.name.toLowerCase());

    if (existingParam && existingParam.id !== formData.id) {
        ui.updateCustomParamStatus('This parameter name is already in use.', 'red');
        return;
    }

    ui.updateCustomParamStatus('Saving...');
    
    if (formData.id) { // Editing existing
        const index = currentParams.findIndex(p => p.id === formData.id);
        if (index > -1) currentParams[index] = { ...currentParams[index], ...formData };
    } else { // Creating new
        formData.id = 'sparam_' + new Date().getTime();
        currentParams.push(formData);
    }
    
    await data.saveSettings(SETTINGS_KEYS.CUSTOM_PARAMS, currentParams);
    updateState('customParameters', currentParams);
    ui.populateParameterButtons();
    
    ui.updateCustomParamStatus('Parameter saved successfully!', 'green');
    setTimeout(() => customParamModal.hide(), 1500);
}

async function handleEditCustomParameter(paramId) {
    const params = await data.getSettings(SETTINGS_KEYS.CUSTOM_PARAMS);
    const param = params.find(p => p.id === paramId);
    if (param) {
        manageCustomParamsModal.hide();
        ui.showCustomParamModal(param);
    }
}

async function handleDeleteCustomParameter(paramId) {
    let params = await data.getSettings(SETTINGS_KEYS.CUSTOM_PARAMS);
    params = params.filter(p => p.id !== paramId);
    await data.saveSettings(SETTINGS_KEYS.CUSTOM_PARAMS, params);
    updateState('customParameters', params);
    ui.populateParameterButtons();
    await handleManageCustomParams(); // Refresh list
}

async function handleDuplicateCustomParameter(paramId) {
    let params = await data.getSettings(SETTINGS_KEYS.CUSTOM_PARAMS);
    const paramToDuplicate = params.find(p => p.id === paramId);
    if (!paramToDuplicate) return;

    const newParam = JSON.parse(JSON.stringify(paramToDuplicate));
    newParam.name = `${newParam.name}Copy`;
    newParam.id = 'sparam_' + new Date().getTime();
    params.push(newParam);

    await data.saveSettings(SETTINGS_KEYS.CUSTOM_PARAMS, params);
    updateState('customParameters', params);
    ui.populateParameterButtons();
    await handleManageCustomParams(); // Refresh list
}

