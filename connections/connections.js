'use strict';

const SETTINGS_KEY = "studentRetentionConnections";

Office.onReady((info) => {
    if (info.host === Office.HostType.Excel) {
        // Wire up event handlers
        document.getElementById("new-connection-button").onclick = showNewConnectionModal;
        document.getElementById("cancel-new-connection-button").onclick = hideNewConnectionModal;
        document.getElementById("insert-key-button").onclick = showInsertKeyModal;
        document.getElementById("custom-key-button").onclick = () => {
            console.log("Custom Key button clicked - no action defined yet.");
        };
        document.getElementById("cancel-upload-button").onclick = hideInsertKeyModal;
        document.getElementById("key-file-input").onchange = handleFileSelect;
        document.getElementById("submit-key-button").onclick = handleSubmitKey;

        // Load and display any existing connections
        loadAndRenderConnections();
    }
});

/**
 * Shows the modal for choosing a new connection type.
 */
function showNewConnectionModal() {
    const modal = document.getElementById("new-connection-modal");
    if (modal) {
        modal.classList.remove('hidden');
    }
}

/**
 * Hides the modal for choosing a new connection type.
 */
function hideNewConnectionModal() {
    const modal = document.getElementById("new-connection-modal");
    if (modal) {
        modal.classList.add('hidden');
    }
}

/**
 * Loads the list of connections from document settings and renders them.
 */
function loadAndRenderConnections() {
    const connectionsString = Office.context.document.settings.get(SETTINGS_KEY);
    let connections = [];
    if (connectionsString) {
        try {
            connections = JSON.parse(connectionsString);
        } catch (e) {
            console.error("Error parsing connections from settings:", e);
            connections = [];
        }
    }
    renderConnections(connections);
}

/**
 * Renders the list of connections to the UI.
 * @param {Array<Object>} connections - An array of connection objects.
 */
function renderConnections(connections) {
    const container = document.getElementById("connections-list-container");
    const noConnectionsMessage = document.getElementById("no-connections-message");
    
    // Clear previous list but keep the placeholder message
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

            connectionElement.appendChild(nameElement);
            connectionElement.appendChild(descriptionElement);
            container.insertBefore(connectionElement, noConnectionsMessage);
        });
    } else {
        noConnectionsMessage.classList.remove('hidden');
    }
}

/**
 * Shows the modal for uploading a key file.
 */
function showInsertKeyModal() {
    hideNewConnectionModal(); // Hide the first modal
    const modal = document.getElementById("insert-key-modal");
    if (modal) {
        modal.classList.remove('hidden');
    }
}

/**
 * Hides the key file upload modal and resets its state.
 */
function hideInsertKeyModal() {
    const modal = document.getElementById("insert-key-modal");
    if (modal) {
        modal.classList.add('hidden');
        resetFileInput();
    }
}

/**
 * Handles the file selection event from the input.
 * @param {Event} event The file input change event.
 */
function handleFileSelect(event) {
    const fileInput = event.target;
    const submitButton = document.getElementById("submit-key-button");
    const fileUploadText = document.getElementById("file-upload-text");
    const fileNameDisplay = document.getElementById("file-name-display");

    if (fileInput.files.length > 0) {
        const file = fileInput.files[0];
        if (file.type === "application/json") {
            fileUploadText.textContent = "File Selected:";
            fileNameDisplay.textContent = file.name;
            submitButton.disabled = false;
        } else {
            fileUploadText.textContent = "Invalid file type. Please select a .json file.";
            fileNameDisplay.textContent = "";
            submitButton.disabled = true;
        }
    } else {
        resetFileInput();
    }
}

/**
 * Resets the file input and associated UI elements to their default state.
 */
function resetFileInput() {
    document.getElementById("key-file-input").value = "";
    document.getElementById("submit-key-button").disabled = true;
    document.getElementById("file-upload-text").textContent = "Click to upload a JSON file";
    document.getElementById("file-name-display").textContent = "";
}

/**
 * Handles the submission of the selected key file.
 * Reads the file, adds it to settings, and refreshes the list.
 */
function handleSubmitKey() {
    const fileInput = document.getElementById("key-file-input");
    if (fileInput.files.length > 0) {
        const file = fileInput.files[0];
        const reader = new FileReader();

        reader.onload = function(event) {
            try {
                const newConnection = JSON.parse(event.target.result);
                
                // --- Save to settings ---
                const connectionsString = Office.context.document.settings.get(SETTINGS_KEY);
                let connections = connectionsString ? JSON.parse(connectionsString) : [];
                connections.push(newConnection);

                Office.context.document.settings.set(SETTINGS_KEY, JSON.stringify(connections));
                Office.context.document.settings.saveAsync((asyncResult) => {
                    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                        console.error("Failed to save settings: " + asyncResult.error.message);
                    } else {
                        console.log("Settings saved successfully with new connection.");
                        renderConnections(connections);
                        hideInsertKeyModal();
                    }
                });

            } catch (e) {
                alert("Error: The selected file is not valid JSON.");
                console.error("JSON parsing error:", e);
                resetFileInput();
            }
        };

        reader.onerror = function() {
            alert("An error occurred while reading the file.");
            resetFileInput();
        };

        reader.readAsText(file);
    }
}

