import { DOM_IDS, SETTINGS_KEYS } from './constants.js';
import { getState, updateState } from './state.js';
import * as ui from './ui.js';
import * as data from './data.js';
import { Modal } from './modal.js';

// --- Application Modals ---
let exampleModal, payloadModal, sendConfirmModal, templatesModal, saveTemplateModal, customParamModal, manageCustomParamsModal;

// --- Main Application Initialization ---

Office.onReady((info) => {
    if (info.host === Office.HostType.Excel) {
        // Initialize UI components first
        ui.initializeQuill();
        ui.setupPillboxInputs();
        
        // Instantiate all modals for the application
        initializeModals();

        // Then, set up all event listeners
        setupEventListeners();

        // Finally, load initial data and set the view
        loadInitialData();
    }
});

/**
 * Creates instances of the Modal class for all modal dialogs.
 */
function initializeModals() {
    exampleModal = new Modal(DOM_IDS.EXAMPLE_MODAL, {
        closeButtonIds: [DOM_IDS.CLOSE_EXAMPLE_MODAL_BUTTON]
    });
    payloadModal = new Modal(DOM_IDS.PAYLOAD_MODAL, {
        closeButtonIds: [DOM_IDS.CLOSE_PAYLOAD_MODAL_BUTTON]
    });
    sendConfirmModal = new Modal(DOM_IDS.SEND_CONFIRM_MODAL, {
        closeButtonIds: [DOM_IDS.CANCEL_SEND_BUTTON]
    });
    templatesModal = new Modal(DOM_IDS.TEMPLATES_MODAL, {
        closeButtonIds: [DOM_IDS.CLOSE_TEMPLATES_MODAL_BUTTON]
    });
    saveTemplateModal = new Modal(DOM_IDS.SAVE_TEMPLATE_MODAL, {
        closeButtonIds: [DOM_IDS.CANCEL_SAVE_TEMPLATE_BUTTON]
    });
    customParamModal = new Modal(DOM_IDS.CUSTOM_PARAM_MODAL, {
        closeButtonIds: [DOM_IDS.CANCEL_CUSTOM_PARAM_BUTTON]
    });
    manageCustomParamsModal = new Modal(DOM_IDS.MANAGE_CUSTOM_PARAMS_MODAL, {
        closeButtonIds: [DOM_IDS.CLOSE_MANAGE_PARAMS_BUTTON]
    });
}


/**
 * Loads initial data from workbook settings and determines the starting view.
 */
async function loadInitialData() {
    try {
        const connection = await data.checkConnection();
        updateState('powerAutomateUrl', connection ? connection.url : null);
        
        const customParams = await data.getSettings(SETTINGS_KEYS.CUSTOM_PARAMS);
        updateState('customParameters', customParams);
        
        ui.populateParameterButtons();

        if (getState().powerAutomateUrl) {
            ui.showView('composer');
        } else {
            ui.showView('setup');
        }
    } catch (error) {
        console.error("Failed to load initial data:", error);
        ui.updateStatus("Error loading initial settings.", "red");
    }
}

/**
 * Attaches all event listeners to the DOM elements.
 */
function setupEventListeners() {
    // Setup Wizard
    document.getElementById(DOM_IDS.CREATE_CONNECTION_BUTTON).onclick = handleCreateConnection;

    // Main Actions
    document.getElementById(DOM_IDS.SEND_EMAIL_BUTTON).onclick = handleSendEmail;
    document.getElementById(DOM_IDS.SHOW_EXAMPLE_BUTTON).onclick = handleShowExample;
    document.getElementById(DOM_IDS.SHOW_PAYLOAD_BUTTON).onclick = handleShowPayload;
    
    // Templates
    document.getElementById(DOM_IDS.TEMPLATES_BUTTON).onclick = handleShowTemplates;
    document.getElementById(DOM_IDS.SAVE_CURRENT_TEMPLATE_BUTTON).onclick = handleShowSaveTemplateModal;
    document.getElementById(DOM_IDS.CONFIRM_SAVE_TEMPLATE_BUTTON).onclick = handleSaveTemplate;

    // Custom Parameters
    document.getElementById(DOM_IDS.CREATE_CUSTOM_PARAM_BUTTON).onclick = () => ui.showCustomParamModal();
    document.getElementById(DOM_IDS.SAVE_CUSTOM_PARAM_BUTTON).onclick = handleSaveCustomParameter;
    document.getElementById(DOM_IDS.ADD_MAPPING_BUTTON).onclick = () => ui.addMappingRow();
    document.getElementById(DOM_IDS.MANAGE_CUSTOM_PARAMS_BUTTON).onclick = handleShowManageCustomParams;


    // Send Confirmation
    document.getElementById(DOM_IDS.CONFIRM_SEND_BUTTON).onclick = executeSend;


    // Dropdowns and Toggles
    document.getElementById(DOM_IDS.RECIPIENT_LIST).onchange = ui.toggleCustomSheetInput;
    
    // Global click listener to close special parameter panels
    document.addEventListener('click', (e) => {
        if (!e.target.closest('.randomize-tag-wrapper') && !e.target.closest('.condition-tag-wrapper')) {
           document.querySelectorAll('.randomize-panel, .condition-panel').forEach(p => p.style.display = 'none');
           document.querySelectorAll('.randomize-arrow, .condition-arrow').forEach(a => a.classList.remove('open'));
       }
   });
}

// --- Event Handlers (Coordinators) ---

async function handleCreateConnection() {
    ui.updateStatus('Creating connection...', 'gray', true);
    const url = ui.getPowerAutomateUrl();

    if (!url.startsWith('http')) {
        ui.updateStatus('Please enter a valid HTTP URL.', 'red', true);
        return;
    }

    try {
        const newConnection = await data.createConnection(url);
        updateState('powerAutomateUrl', newConnection.url);
        ui.updateStatus('Connection created successfully!', 'green', true);
        setTimeout(() => ui.showView('composer'), 1500);
    } catch (error) {
        console.error("Failed to create connection:", error);
        ui.updateStatus('Failed to save the connection.', 'red', true);
    }
}

async function handleShowPayload() {
    ui.updateStatus('Generating payload preview...', 'gray');
    try {
        const sheetName = ui.getSelectedSheetName();
        if (!sheetName) {
            ui.updateStatus('Please select a valid student list or enter a custom sheet name.', 'orange');
            return;
        }
        const { customParameters } = getState();
        const students = await data.getStudentData(sheetName, customParameters);
        updateState('students', students);

        if (students.length === 0) {
            ui.updateStatus('No students found to generate a payload.', 'orange');
            setTimeout(() => ui.updateStatus(''), 3000);
            return;
        }

        const template = ui.getEmailTemplateFromDOM();
        const payload = data.buildPayload(students, template);

        ui.populatePayloadModal(payload);
        payloadModal.show();
        ui.updateStatus('');

    } catch (error) {
        console.error("Error showing payload:", error);
        ui.updateStatus(error.message, 'red');
    }
}

async function handleShowExample() {
    ui.updateStatus('Generating example...', 'gray');
    try {
        const sheetName = ui.getSelectedSheetName();
        if (!sheetName) {
            ui.updateStatus('Please select a valid student list or enter a custom sheet name.', 'orange');
            return;
        }
        const { customParameters } = getState();
        const students = await data.getStudentData(sheetName, customParameters);
        
        if (students.length === 0) {
            ui.updateStatus('No students found to generate an example.', 'orange');
            setTimeout(() => ui.updateStatus(''), 3000);
            return;
        }

        const randomStudent = students[Math.floor(Math.random() * students.length)];
        const template = ui.getEmailTemplateFromDOM();
        
        const examplePayload = data.buildPayload([randomStudent], template)[0];

        // FIX: Check if the payload is valid before trying to show the modal
        if (examplePayload) {
            ui.populateExampleModal(examplePayload);
            exampleModal.show();
            ui.updateStatus('');
        } else {
            ui.updateStatus('Selected example student has no email. Try again.', 'orange');
            setTimeout(() => ui.updateStatus(''), 3000);
        }

    } catch(error) {
        console.error("Error showing example:", error);
        ui.updateStatus(error.message, 'red');
    }
}

async function handleSendEmail() {
    ui.updateStatus('Preparing to send emails...', 'gray');
    try {
        const sheetName = ui.getSelectedSheetName();
        if (!sheetName) {
            ui.updateStatus('Please select a valid student list or enter a custom sheet name.', 'orange');
            return;
        }
        const { customParameters } = getState();
        const students = await data.getStudentData(sheetName, customParameters);
        updateState('students', students);

        if (students.length === 0) {
            ui.updateStatus('No students to send emails to.', 'orange');
            return;
        }
        
        const template = ui.getEmailTemplateFromDOM();
        const payload = data.buildPayload(students, template);
        
        if (payload.length === 0) {
            ui.updateStatus('No students with valid "To" and "From" emails found.', 'orange');
            return;
        }

        ui.populateSendConfirmationModal(payload.length);
        sendConfirmModal.show();

    } catch (error) {
        console.error("Error sending emails:", error);
        ui.updateStatus(`Failed to send emails: ${error.message}`, 'red');
    }
}

async function executeSend() {
    sendConfirmModal.hide();
    const { students, powerAutomateUrl } = getState();
    const template = ui.getEmailTemplateFromDOM();
    const payload = data.buildPayload(students, template);

    ui.updateStatus(`Sending ${payload.length} emails...`, 'gray');
    try {
        await data.sendToPowerAutomate(powerAutomateUrl, payload);
        ui.updateStatus(`Successfully sent ${payload.length} emails!`, 'green');
    } catch(error) {
        console.error("Error sending emails:", error);
        ui.updateStatus(`Failed to send emails: ${error.message}`, 'red');
    }
}

// --- Template Handlers ---

async function handleShowTemplates() {
    const templates = await data.getSettings(SETTINGS_KEYS.EMAIL_TEMPLATES);
    ui.populateTemplatesModal(templates, {
        onLoad: handleLoadTemplate,
        onDelete: handleDeleteTemplate
    });
    templatesModal.show();
}

function handleShowSaveTemplateModal() {
    templatesModal.hide();
    ui.clearSaveTemplateForm();
    saveTemplateModal.show();
}

async function handleSaveTemplate() {
    const { name, author } = ui.getTemplateSaveForm();
    if (!name || !author) {
        ui.updateSaveTemplateStatus('Name and Author are required.', 'red');
        return;
    }
    ui.updateSaveTemplateStatus('Saving...', 'gray');

    const newTemplate = {
        id: 'template_' + new Date().getTime(),
        name,
        author,
        timestamp: new Date().toISOString(),
        ...ui.getEmailTemplateFromDOM()
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
    await handleShowTemplates(); // Refresh the list
}

// --- Custom Parameter Handlers ---

async function handleSaveCustomParameter() {
    const { name, sourceColumn, defaultValue, mappings, id } = ui.getCustomParamForm();
    
    if (!/^[a-zA-Z0-9]+$/.test(name)) {
        ui.updateCustomParamStatus('Name must be alphanumeric with no spaces.', 'red');
        return;
    }
    const { customParameters } = getState();
    const existingParam = customParameters.find(p => p.name.toLowerCase() === name.toLowerCase());
    if (STANDARD_PARAMETERS.includes(name) || (existingParam && existingParam.id !== id)) {
        ui.updateCustomParamStatus('This parameter name is already in use.', 'red');
        return;
    }
    
    ui.updateCustomParamStatus('Saving...', 'gray');

    const paramData = { name, sourceColumn, defaultValue, mappings };
    let currentParams = [...customParameters];

    if (id) {
        const index = currentParams.findIndex(p => p.id === id);
        if (index > -1) currentParams[index] = { ...currentParams[index], ...paramData };
    } else {
        paramData.id = 'sparam_' + new Date().getTime();
        currentParams.push(paramData);
    }
    
    await data.saveSettings(SETTINGS_KEYS.CUSTOM_PARAMS, currentParams);
    updateState('customParameters', currentParams);
    ui.populateParameterButtons();

    ui.updateCustomParamStatus('Parameter saved!', 'green');
    setTimeout(() => customParamModal.hide(), 1500);
}

async function handleShowManageCustomParams() {
    customParamModal.hide();
    const params = await data.getSettings(SETTINGS_KEYS.CUSTOM_PARAMS);
    ui.populateManageCustomParamsModal(params, {
        onEdit: (paramId) => {
            manageCustomParamsModal.hide();
            const param = params.find(p => p.id === paramId);
            ui.showCustomParamModal(param);
        },
        onDelete: async (paramId) => {
            let currentParams = await data.getSettings(SETTINGS_KEYS.CUSTOM_PARAMS);
            currentParams = currentParams.filter(p => p.id !== paramId);
            await data.saveSettings(SETTINGS_KEYS.CUSTOM_PARAMS, currentParams);
            updateState('customParameters', currentParams);
            ui.populateParameterButtons();
            handleShowManageCustomParams(); // Refresh
        },
        onDuplicate: async (paramId) => {
            let params = await data.getSettings(SETTINGS_KEYS.CUSTOM_PARAMS);
            const paramToDuplicate = params.find(p => p.id === paramId);
            if (!paramToDuplicate) return;

            const newParam = JSON.parse(JSON.stringify(paramToDuplicate));
            let newName = `${newParam.name}Copy`;
            const allParamNames = [...STANDARD_PARAMETERS, ...params.map(p => p.name)];
            while (allParamNames.includes(newName)) newName = `${newName}Copy`;
            newParam.name = newName;
            newParam.id = 'sparam_' + new Date().getTime();
            params.push(newParam);
            await data.saveSettings(SETTINGS_KEYS.CUSTOM_PARAMS, params);
            updateState('customParameters', params);
            ui.populateParameterButtons();
            handleShowManageCustomParams(); // Refresh
        }
    });
    manageCustomParamsModal.show();
}

