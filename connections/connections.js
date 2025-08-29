'use strict';

const SETTINGS_KEY = "studentRetentionConnections";
let wizardState = {};
let connectionToDeleteIndex = null;


Office.onReady((info) => {
    if (info.host === Office.HostType.Excel) {
        document.getElementById("new-connection-button").onclick = showNewConnectionModal;
        document.getElementById("cancel-new-connection-button").onclick = hideNewConnectionModal;
        document.getElementById("insert-key-button").onclick = showInsertKeyModal;
        document.getElementById("custom-key-button").onclick = showCustomKeyWizard;
        document.getElementById("cancel-upload-button").onclick = hideInsertKeyModal;
        document.getElementById("key-file-input").onchange = handleFileSelect;
        document.getElementById("submit-key-button").onclick = handleSubmitKey;
        document.getElementById("wizard-cancel-button").onclick = hideCustomKeyWizard;
        document.getElementById("wizard-back-button").onclick = navigateWizardBack;
        document.getElementById("wizard-next-button").onclick = navigateWizardNext;
        document.getElementById("wizard-finish-button").onclick = finishCustomKeyWizard;
        document.getElementById("pusher-setup-button").onclick = handlePusherSelect;
        document.getElementById("manual-setup-button").onclick = handleManualSelect;
        document.querySelectorAll(".wizard-type-button").forEach(btn => {
            btn.addEventListener('click', (e) => selectConnectionType(e.currentTarget));
        });
        document.getElementById("cancel-delete-button").onclick = hideDeleteConfirmModal;
        document.getElementById("confirm-delete-button").onclick = handleDeleteConnection;

        loadAndRenderConnections();
    }
});

function showNewConnectionModal() {
    document.getElementById("new-connection-modal").classList.remove('hidden');
}

function hideNewConnectionModal() {
    document.getElementById("new-connection-modal").classList.add('hidden');
}

function loadAndRenderConnections() {
    const connectionsString = Office.context.document.settings.get(SETTINGS_KEY);
    let connections = connectionsString ? JSON.parse(connectionsString) : [];
    renderConnections(connections);
}

function renderConnections(connections) {
    const container = document.getElementById("connections-list-container");
    const noConnectionsMessage = document.getElementById("no-connections-message");
    
    container.querySelectorAll(".connection-item").forEach(item => item.remove());

    if (connections && connections.length > 0) {
        noConnectionsMessage.classList.add('hidden');
        connections.forEach((conn, index) => {
            const connectionElement = document.createElement('div');
            connectionElement.className = "connection-item bg-white p-4 rounded-lg shadow-sm border border-gray-200 relative group";
            connectionElement.dataset.connectionIndex = index;

            const optionsContainer = document.createElement('div');
            optionsContainer.className = "absolute top-2 right-2";

            const optionsButton = document.createElement('button');
            optionsButton.className = "options-button opacity-0 group-hover:opacity-100 p-1 rounded-full hover:bg-gray-200 transition-opacity focus:opacity-100";
            optionsButton.innerHTML = `<svg class="w-5 h-5 text-gray-600" fill="none" stroke="currentColor" viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M12 5v.01M12 12v.01M12 19v.01M12 6a1 1 0 110-2 1 1 0 010 2zm0 7a1 1 0 110-2 1 1 0 010 2zm0 7a1 1 0 110-2 1 1 0 010 2z"></path></svg>`;
            
            const optionsMenu = document.createElement('div');
            optionsMenu.className = "options-menu hidden absolute right-0 mt-2 w-32 bg-white rounded-md shadow-lg z-10 border";
            
            const editLink = document.createElement('a');
            editLink.href = "#";
            editLink.className = "edit-btn block px-4 py-2 text-sm text-gray-700 hover:bg-gray-100";
            editLink.textContent = "Edit";
            editLink.onclick = (e) => { e.preventDefault(); handleEditConnection(index); };

            const deleteLink = document.createElement('a');
            deleteLink.href = "#";
            deleteLink.className = "delete-btn block px-4 py-2 text-sm text-red-600 hover:bg-gray-100";
            deleteLink.textContent = "Delete";
            deleteLink.onclick = (e) => { e.preventDefault(); promptDeleteConnection(index); };
            
            optionsMenu.appendChild(editLink);
            optionsMenu.appendChild(deleteLink);
            optionsContainer.appendChild(optionsButton);
            optionsContainer.appendChild(optionsMenu);
            connectionElement.appendChild(optionsContainer);
            
            optionsButton.onclick = (e) => {
                e.stopPropagation();
                document.querySelectorAll('.options-menu').forEach(menu => {
                    if (menu !== optionsMenu) menu.classList.add('hidden');
                });
                optionsMenu.classList.toggle('hidden');
            };

            const nameElement = document.createElement('h3');
            nameElement.className = "font-bold text-md text-gray-800 pr-8";
            nameElement.textContent = conn.name || "Unnamed Connection";
            
            const descriptionElement = document.createElement('p');
            descriptionElement.className = "text-sm text-gray-600 mt-1";
            descriptionElement.textContent = conn.description || "No description provided.";
            
            connectionElement.appendChild(nameElement);
            connectionElement.appendChild(descriptionElement);

            if(conn.type) {
                const typeElement = document.createElement('span');
                let providerText = conn.credentials && conn.credentials.provider ? conn.credentials.provider : conn.type;
                typeElement.className = "text-xs font-semibold inline-block py-1 px-2 uppercase rounded-full text-blue-600 bg-blue-200 uppercase mt-2";
                typeElement.textContent = providerText;
                connectionElement.appendChild(typeElement);
            }

            container.insertBefore(connectionElement, noConnectionsMessage);
        });
    } else {
        noConnectionsMessage.classList.remove('hidden');
    }
}

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

function showCustomKeyWizard() {
    hideNewConnectionModal();
    resetWizard();
    document.getElementById("custom-key-wizard-modal").classList.remove('hidden');
}

function hideCustomKeyWizard() {
    document.getElementById("custom-key-wizard-modal").classList.add('hidden');
}

function resetWizard() {
    wizardState = { step: 'setupType' };
    
    // Hide all steps
    document.getElementById('wizard-step-1-setup-type').classList.remove('hidden');
    document.getElementById('wizard-step-1-manual-type').classList.add('hidden');
    document.getElementById('wizard-step-details').classList.add('hidden');
    document.getElementById('wizard-step-pusher-creds').classList.add('hidden');
    
    // Reset buttons
    document.getElementById('wizard-back-button').classList.add('hidden');
    document.getElementById('wizard-finish-button').classList.add('hidden');
    const nextButton = document.getElementById('wizard-next-button');
    nextButton.classList.add('hidden');
    nextButton.disabled = true;

    // Clear selections and inputs
    document.querySelectorAll(".wizard-type-button, .wizard-setup-button").forEach(btn => {
        btn.classList.remove('bg-blue-100', 'border-blue-400');
    });
    document.getElementById('connection-name').value = '';
    document.getElementById('connection-description').value = '';
    document.getElementById('pusher-app-id').value = '';
    document.getElementById('pusher-key').value = '';
    document.getElementById('pusher-secret').value = '';
    document.getElementById('pusher-cluster').value = '';
}

function handlePusherSelect() {
    wizardState = { ...wizardState, step: 'details', setupType: 'Pusher', type: 'Webhook' };
    updateWizardView();
}

function handleManualSelect() {
    wizardState = { ...wizardState, step: 'manualType' };
    updateWizardView();
}

function selectConnectionType(buttonElement) {
    wizardState.type = buttonElement.dataset.type;
    document.querySelectorAll(".wizard-type-button").forEach(btn => btn.classList.remove('bg-blue-100', 'border-blue-400'));
    buttonElement.classList.add('bg-blue-100', 'border-blue-400');
    document.getElementById('wizard-next-button').disabled = false;
}


function navigateWizardNext() {
    if (wizardState.step === 'manualType') {
        wizardState.step = 'details';
    } else if (wizardState.step === 'details' && wizardState.type === 'Webhook') {
        wizardState.step = 'pusherCreds';
    }
    updateWizardView();
}

function navigateWizardBack() {
    if (wizardState.step === 'details' && wizardState.setupType === 'Pusher') {
        wizardState.step = 'setupType';
    } else if (wizardState.step === 'details' && wizardState.setupType !== 'Pusher') {
        wizardState.step = 'manualType';
    } else if (wizardState.step === 'manualType') {
        wizardState.step = 'setupType';
    } else if (wizardState.step === 'pusherCreds') {
        wizardState.step = 'details';
    }
    updateWizardView();
}

function updateWizardView() {
    const { step, type, setupType } = wizardState;
    
    // Hide all steps first
    document.getElementById('wizard-step-1-setup-type').classList.add('hidden');
    document.getElementById('wizard-step-1-manual-type').classList.add('hidden');
    document.getElementById('wizard-step-details').classList.add('hidden');
    document.getElementById('wizard-step-pusher-creds').classList.add('hidden');

    const nextButton = document.getElementById('wizard-next-button');
    const backButton = document.getElementById('wizard-back-button');
    const finishButton = document.getElementById('wizard-finish-button');

    // Show current step
    if (step === 'setupType') {
        document.getElementById('wizard-step-1-setup-type').classList.remove('hidden');
        backButton.classList.add('hidden');
        nextButton.classList.add('hidden');
        finishButton.classList.add('hidden');
    } else if (step === 'manualType') {
        document.getElementById('wizard-step-1-manual-type').classList.remove('hidden');
        backButton.classList.remove('hidden');
        nextButton.classList.remove('hidden');
        nextButton.disabled = !type; // Enable if a type is already selected
        finishButton.classList.add('hidden');
    } else if (step === 'details') {
        document.getElementById('wizard-step-details').classList.remove('hidden');
        document.getElementById('connection-type-display').textContent = setupType === 'Pusher' ? 'Pusher Webhook' : type;
        backButton.classList.remove('hidden');
        if (type === 'Webhook') {
            nextButton.classList.remove('hidden');
            nextButton.disabled = false;
            finishButton.classList.add('hidden');
        } else {
            nextButton.classList.add('hidden');
            finishButton.classList.remove('hidden');
        }
    } else if (step === 'pusherCreds') {
        document.getElementById('wizard-step-pusher-creds').classList.remove('hidden');
        backButton.classList.remove('hidden');
        nextButton.classList.add('hidden');
        finishButton.classList.remove('hidden');
    }
}


function finishCustomKeyWizard() {
    const name = document.getElementById('connection-name').value.trim();
    const description = document.getElementById('connection-description').value.trim();
    
    if (!name) {
        alert("Please enter a connection name.");
        return;
    }

    const newConnection = {
        type: wizardState.type,
        name: name,
        description: description
    };
    
    if (wizardState.type === 'Webhook') {
        const appId = document.getElementById('pusher-app-id').value.trim();
        const key = document.getElementById('pusher-key').value.trim();
        const secret = document.getElementById('pusher-secret').value.trim();
        const cluster = document.getElementById('pusher-cluster').value.trim();

        if (!appId || !key || !secret || !cluster) {
            alert("Please fill out all Pusher credential fields.");
            return;
        }

        newConnection.credentials = {
            provider: 'Pusher',
            appId: appId,
            key: key,
            secret: secret,
            cluster: cluster
        };
    }

    saveConnection(newConnection);
    hideCustomKeyWizard();
}

function saveConnection(newConnection) {
    const connectionsString = Office.context.document.settings.get(SETTINGS_KEY);
    let connections = connectionsString ? JSON.parse(connectionsString) : [];
    connections.push(newConnection);
    saveConnectionsArray(connections);
}

function saveConnectionsArray(connections) {
     Office.context.document.settings.set(SETTINGS_KEY, JSON.stringify(connections));
    Office.context.document.settings.saveAsync((asyncResult) => {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
            console.error("Failed to save settings: " + asyncResult.error.message);
            alert("Error: Could not save the connections.");
        } else {
            console.log("Settings saved successfully.");
            renderConnections(connections);
        }
    });
}

function handleEditConnection(index) {
    console.log("Editing connection at index:", index);
    alert("Editing functionality is not yet implemented.");
}

function promptDeleteConnection(index) {
    connectionToDeleteIndex = index;
    document.getElementById("delete-confirm-modal").classList.remove('hidden');
}

function hideDeleteConfirmModal() {
    document.getElementById("delete-confirm-modal").classList.add('hidden');
    connectionToDeleteIndex = null;
}

function handleDeleteConnection() {
    if (connectionToDeleteIndex !== null) {
        const connectionsString = Office.context.document.settings.get(SETTINGS_KEY);
        let connections = connectionsString ? JSON.parse(connectionsString) : [];
        connections.splice(connectionToDeleteIndex, 1);
        saveConnectionsArray(connections);
        hideDeleteConfirmModal();
    }
}

document.addEventListener('click', (event) => {
    document.querySelectorAll('.options-menu').forEach(menu => {
        const button = menu.previousElementSibling;
        if (!menu.contains(event.target) && !button.contains(event.target)) {
            menu.classList.add('hidden');
        }
    });
});

