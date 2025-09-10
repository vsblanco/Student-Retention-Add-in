import { DOM_IDS, SETTINGS_KEYS } from './constants.js';
import { getState, updateState } from './state.js';
import * as ui from './ui.js';
import * as data from './data.js';

// --- Main Application Initialization ---

Office.onReady((info) => {
    if (info.host === Office.HostType.Excel) {
        // Initialize UI components first
        ui.initializeQuill();
        ui.setupPillboxInputs();
        
        // Then, set up all event listeners
        setupEventListeners();

        // Finally, load initial data and set the view
        loadInitialData();
    }
});

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

    // Dropdowns and Toggles
    document.getElementById(DOM_IDS.RECIPIENT_LIST).onchange = ui.toggleCustomSheetInput;
    
    // Modal Close Buttons
    document.getElementById(DOM_IDS.CLOSE_EXAMPLE_MODAL_BUTTON).onclick = () => ui.hideModal(DOM_IDS.EXAMPLE_MODAL);
    document.getElementById(DOM_IDS.CLOSE_PAYLOAD_MODAL_BUTTON).onclick = () => ui.hideModal(DOM_IDS.PAYLOAD_MODAL);

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

        ui.displayPayloadInModal(payload);
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

        ui.displayExampleInModal(examplePayload);
        ui.updateStatus('');

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
        const { customParameters, powerAutomateUrl } = getState();
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

        const userConfirmed = await ui.showSendConfirmation(payload.length);
        
        if (userConfirmed) {
            ui.updateStatus(`Sending ${payload.length} emails...`, 'gray');
            await data.sendToPowerAutomate(powerAutomateUrl, payload);
            ui.updateStatus(`Successfully sent ${payload.length} emails!`, 'green');
        } else {
            ui.updateStatus('Send cancelled.', 'gray');
        }
    } catch (error) {
        console.error("Error sending emails:", error);
        ui.updateStatus(`Failed to send emails: ${error.message}`, 'red');
    }
}
