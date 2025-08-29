'use strict';

const SETTINGS_KEY = "studentRetentionConnections";
let wizardState = {
    step: 1,
    type: null,
    name: '',
    description: ''
};


Office.onReady((info) => {
    if (info.host === Office.HostType.Excel) {
        // Wire up event handlers
        document.getElementById("new-connection-button").onclick = showNewConnectionModal;
        document.getElementById("cancel-new-connection-button").onclick = hideNewConnectionModal;
        
        document.getElementById("insert-key-button").onclick = showInsertKeyModal;
        document.getElementById("custom-key-button").onclick = showCustomKeyWizard;

        // File Upload Modal Handlers
        document.getElementById("cancel-upload-button").onclick = hideInsertKeyModal;
        document.getElementById("key-file-input").onchange = handleFileSelect;
        document.getElementById("submit-key-button").onclick = handleSubmitKey;

        // Custom Key Wizard Handlers
        document.getElementById("wizard-cancel-button").onclick = hideCustomKeyWizard;
        document.getElementById("wizard-back-button").onclick = navigateWizardBack;
        document.getElementById("wizard-next-button").onclick = navigateWizardNext;
        document.getElementById("wizard-finish-button").onclick = finishCustomKeyWizard;
        document.querySelectorAll(".wizard-type-button").forEach(btn => {
            btn.addEventListener('click', (e) => selectConnectionType(e.currentTarget));
        });

        // Load and display any existing connections
        loadAndRenderConnections();
    }
});

// --- Main Modal Logic ---
function showNewConnectionModal() {
    document.getElementById("new-connection-modal").classList.remove('hidden');
}

function hideNewConnectionModal() {
    document.getElementById("new-connection-modal").classList.add('hidden');
}

// --- Connection Loading and Rendering ---
function loadAndRenderConnections() {
    const connectionsString = Office.context.document.settings.get(SETTINGS_KEY);
    let connections = connectionsString ? JSON.parse(connectionsString) : [];
    renderConnections(connections);
}

function renderConnections(connections) {
    const container = document.getElementById("connections-list-container");
    const noConnectionsMessage = document.getElementById("no-connections-message");
    
    while (container.firstChild && container.firstChild !== noConnectionsMessage) {
        container.removeChild(container.firstChild);
    }

    if (connections && connections.length > 0) {
        noConnectionsMessage.classList.add('hidden');
        connections.forEach(conn => {
            const connectionElement = document.createElement('div');
            connectionElement.className = "bg-white p-4 rounded-lg shadow-sm border border-gray-200";
            
            const nameElement = document.createElement('h3');
            nameElement.className = "font-bold text-md text-gray-800";
            nameElement.textContent = conn.name || "Unnamed Connection";
            
            const descriptionElement = document.createElement('p');
            descriptionElement.className = "text-sm text-gray-600 mt-1";
            descriptionElement.textContent = conn.description || "No description provided.";
            
            if(conn.type) {
                const typeElement = document.createElement('span');
                typeElement.className = "text-xs font-semibold inline-block py-1 px-2 uppercase rounded-full text-blue-600 bg-blue-200 uppercase last:mr-0 mr-1 mt-2";
                typeElement.textContent = conn.type;
                connectionElement.appendChild(typeElement);
            }

            connectionElement.insertBefore(nameElement, connectionElement.firstChild);
            connectionElement.appendChild(descriptionElement);

            container.insertBefore(connectionElement, noConnectionsMessage);
        });
    } else {
        noConnectionsMessage.classList.remove('hidden');
    }
}

// --- Insert Key (File Upload) Modal Logic ---
function showInsertKeyModal() {
    hideNewConnectionModal(); 
    document.getElementById("insert-key-modal").classList.remove('hidden');
}

function hideInsertKeyModal() {
    document.getElementById("insert-key-modal").classList.add('hidden');
    resetFileInput();
}

function handleFileSelect(event) {
    const fileInput = event.target;
    const submitButton = document.getElementById("submit-key-button");
    const fileNameDisplay = document.getElementById("file-name-display");

    if (fileInput.files.length > 0) {
        const file = fileInput.files[0];
        if (file.type === "application/json") {
            fileNameDisplay.textContent = file.name;
            submitButton.disabled = false;
        } else {
            fileNameDisplay.textContent = "Invalid file type. Must be .json";
            submitButton.disabled = true;
        }
    }
}

function resetFileInput() {
    document.getElementById("key-file-input").value = "";
    document.getElementById("submit-key-button").disabled = true;
    document.getElementById("file-name-display").textContent = "";
}

function handleSubmitKey() {
    const fileInput = document.getElementById("key-file-input");
    if (fileInput.files.length > 0) {
        const file = fileInput.files[0];
        const reader = new FileReader();
        reader.onload = (event) => {
            try {
                const newConnection = JSON.parse(event.target.result);
                saveConnection(newConnection);
                hideInsertKeyModal();
            } catch (e) {
                alert("Error: The selected file is not valid JSON.");
                resetFileInput();
            }
        };
        reader.readAsText(file);
    }
}


// --- Custom Key Wizard Logic ---
function showCustomKeyWizard() {
    hideNewConnectionModal();
    resetWizard();
    document.getElementById("custom-key-wizard-modal").classList.remove('hidden');
}

function hideCustomKeyWizard() {
    document.getElementById("custom-key-wizard-modal").classList.add('hidden');
}

function resetWizard() {
    wizardState = { step: 1, type: null, name: '', description: '' };
    
    document.getElementById('wizard-step-1').classList.remove('hidden');
    document.getElementById('wizard-step-2').classList.add('hidden');
    document.getElementById('wizard-step-3').classList.add('hidden');
    
    document.getElementById('wizard-back-button').classList.add('hidden');
    document.getElementById('wizard-finish-button').classList.add('hidden');
    document.getElementById('wizard-next-button').classList.remove('hidden');
    document.getElementById('wizard-next-button').disabled = true;

    document.querySelectorAll(".wizard-type-button").forEach(btn => {
        btn.classList.remove('bg-blue-100', 'border-blue-400');
    });

    document.getElementById('connection-name').value = '';
    document.getElementById('connection-description').value = '';
    document.getElementById('pusher-app-id').value = '';
    document.getElementById('pusher-key').value = '';
    document.getElementById('pusher-secret').value = '';
    document.getElementById('pusher-cluster').value = '';
}

function selectConnectionType(buttonElement) {
    document.querySelectorAll(".wizard-type-button").forEach(btn => {
        btn.classList.remove('bg-blue-100', 'border-blue-400');
    });
    buttonElement.classList.add('bg-blue-100', 'border-blue-400');
    wizardState.type = buttonElement.dataset.type;
    document.getElementById('wizard-next-button').disabled = false;
}

function navigateWizardNext() {
    if (wizardState.step === 1) {
        if (wizardState.type === 'Webhook') {
            // For Webhook, go to the special credentials step
            wizardState.step = 3;
            document.getElementById('wizard-step-1').classList.add('hidden');
            document.getElementById('wizard-step-3').classList.remove('hidden');

            document.getElementById('wizard-back-button').classList.remove('hidden');
            document.getElementById('wizard-next-button').classList.add('hidden');
            document.getElementById('wizard-finish-button').classList.remove('hidden');

        } else {
            // For API and HTTP, go to the standard details step
            wizardState.step = 2;
            document.getElementById('wizard-step-1').classList.add('hidden');
            document.getElementById('wizard-step-2').classList.remove('hidden');
            
            document.getElementById('wizard-back-button').classList.remove('hidden');
            document.getElementById('wizard-next-button').classList.add('hidden');
            document.getElementById('wizard-finish-button').classList.remove('hidden');
    
            document.getElementById('connection-type-display').textContent = wizardState.type;
        }
    }
}

function navigateWizardBack() {
    if (wizardState.step === 2 || wizardState.step === 3) {
        wizardState.step = 1;
        document.getElementById('wizard-step-2').classList.add('hidden');
        document.getElementById('wizard-step-3').classList.add('hidden');
        document.getElementById('wizard-step-1').classList.remove('hidden');

        document.getElementById('wizard-back-button').classList.add('hidden');
        document.getElementById('wizard-finish-button').classList.add('hidden');
        document.getElementById('wizard-next-button').classList.remove('hidden');
        document.getElementById('wizard-next-button').disabled = false; // It was selected before
    }
}

function finishCustomKeyWizard() {
    let newConnection;
    if (wizardState.type === 'Webhook') {
        // Collect all details for Webhook
        const name = document.getElementById('connection-name').value.trim(); // Name is on step 2, but we need it now
        const description = document.getElementById('connection-description').value.trim(); // Also step 2
        
        const appId = document.getElementById('pusher-app-id').value.trim();
        const key = document.getElementById('pusher-key').value.trim();
        const secret = document.getElementById('pusher-secret').value.trim();
        const cluster = document.getElementById('pusher-cluster').value.trim();

        if (!appId || !key || !secret || !cluster) {
            alert("Please fill out all Pusher credential fields.");
            return;
        }
        
        // Let's go to step 2 to get name/description first
        wizardState.step = 2;
        document.getElementById('wizard-step-3').classList.add('hidden');
        document.getElementById('wizard-step-2').classList.remove('hidden');
        document.getElementById('connection-type-display').textContent = wizardState.type;
        return; // Wait for user to fill details

    } 
    
    // This part runs after step 2 is filled out
    const name = document.getElementById('connection-name').value.trim();
    const description = document.getElementById('connection-description').value.trim();
    
    if (!name) {
        alert("Please enter a connection name.");
        return;
    }

    newConnection = {
        type: wizardState.type,
        name: name,
        description: description
    };
    
    if(wizardState.type === 'Webhook') {
        newConnection.credentials = {
            provider: 'Pusher',
            appId: document.getElementById('pusher-app-id').value.trim(),
            key: document.getElementById('pusher-key').value.trim(),
            secret: document.getElementById('pusher-secret').value.trim(),
            cluster: document.getElementById('pusher-cluster').value.trim()
        };
    }

    saveConnection(newConnection);
    hideCustomKeyWizard();
}

// --- Universal Save Function ---
function saveConnection(newConnection) {
    const connectionsString = Office.context.document.settings.get(SETTINGS_KEY);
    let connections = connectionsString ? JSON.parse(connectionsString) : [];
    connections.push(newConnection);

    Office.context.document.settings.set(SETTINGS_KEY, JSON.stringify(connections));
    Office.context.document.settings.saveAsync((asyncResult) => {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
            console.error("Failed to save settings: " + asyncResult.error.message);
            alert("Error: Could not save the new connection.");
        } else {
            console.log("Settings saved successfully with new connection.");
            renderConnections(connections);
        }
    });
}

