'use strict';

// --- STATE & CONSTANTS ---
const SETTINGS_KEY = "studentRetentionSettings";
const MASTER_LIST_SHEET = "Master List";

let activePusherInstances = {}; // Stores { connectionId: { pusher, channel } }

// --- INITIALIZATION ---
Office.onReady((info) => {
    if (info.host === Office.HostType.Excel) {
        initializeEventListeners();
        logToUI("Connections pane ready.");
        loadAndRenderConnections();
    }
});

function initializeEventListeners() {
    // Main UI
    document.getElementById("new-connection-button").onclick = showSelectServiceModal;
    
    // Select Service Modal
    document.getElementById("cancel-select-service-button").onclick = hideSelectServiceModal;
    document.getElementById("select-pusher-button").onclick = showPusherConfigModal;
    
    // Pusher Config Modal
    document.getElementById("cancel-pusher-config-button").onclick = hidePusherConfigModal;
    document.getElementById("create-pusher-connection-button").onclick = handleCreatePusherConnection;
    
    // Log panel
    document.getElementById("log-header").onclick = toggleLogPanel;
    document.getElementById("clear-log-button").onclick = clearLogPanel;

    // Event delegation for connection cards
    document.getElementById("connections-list-container").addEventListener('click', handleConnectionCardAction);
}

// --- LOGGING ---
function logToUI(message, type = 'INFO') {
    const logContainer = document.getElementById('log-container');
    if (logContainer) {
        const p = document.createElement('p');
        const timestamp = new Date().toLocaleTimeString();
        p.innerHTML = `<span class="text-gray-500">${timestamp} [${type}]:</span> ${message}`;
        
        switch (type) {
            case 'ERROR': p.className = 'text-red-400'; break;
            case 'SUCCESS': p.className = 'text-green-400'; break;
            default: p.className = 'text-gray-300'; break;
        }
        logContainer.appendChild(p);
        logContainer.scrollTop = logContainer.scrollHeight;
    }
    console.log(`[${type}] ${message}`);
}

function toggleLogPanel() {
    document.getElementById('log-container').classList.toggle('hidden');
    document.getElementById('log-arrow').style.transform = document.getElementById('log-container').classList.contains('hidden') ? 'rotate(0deg)' : 'rotate(180deg)';
}

function clearLogPanel(e) {
    if (e) e.stopPropagation();
    document.getElementById('log-container').innerHTML = '';
    logToUI("Log cleared.");
}

// --- MODAL & UI MANAGEMENT ---
function showSelectServiceModal() { document.getElementById("select-service-modal").classList.remove('hidden'); }
function hideSelectServiceModal() { document.getElementById("select-service-modal").classList.add('hidden'); }
function showPusherConfigModal() { hideSelectServiceModal(); document.getElementById("pusher-config-modal").classList.remove('hidden'); }
function hidePusherConfigModal() { document.getElementById("pusher-config-modal").classList.add('hidden'); resetPusherForm(); }

function resetPusherForm() {
    document.getElementById("connection-name").value = '';
    document.getElementById("pusher-key").value = '';
    document.getElementById("pusher-secret").value = '';
    document.getElementById("pusher-cluster").value = '';
    document.getElementById("pusher-channel").value = '';
    document.getElementById("pusher-event").value = '';
}

// --- SETTINGS MANAGEMENT ---
async function getSettings() {
    await new Promise(resolve => Office.context.document.settings.refreshAsync(() => resolve()));
    const settingsString = Office.context.document.settings.get(SETTINGS_KEY);
    const defaults = { connections: [] };
    if (settingsString) {
        try {
            const settings = JSON.parse(settingsString);
            return { ...defaults, ...settings };
        } catch (e) {
            logToUI("Error parsing settings, returning defaults.", "ERROR");
            return defaults;
        }
    }
    return defaults;
}

async function saveSettings(settingsToSave) {
    Office.context.document.settings.set(SETTINGS_KEY, JSON.stringify(settingsToSave));
    await new Promise((resolve, reject) => {
        Office.context.document.settings.saveAsync(result => {
            if (result.status === Office.AsyncResultStatus.Failed) {
                reject(new Error("Failed to save settings: " + result.error.message));
            } else {
                logToUI("Settings saved successfully.");
                resolve();
            }
        });
    });
}

// --- CONNECTION MANAGEMENT ---
async function loadAndRenderConnections() {
    logToUI("Loading all connections...");
    const settings = await getSettings();
    renderConnections(settings.connections);
    settings.connections.forEach(conn => {
        if (conn.type === 'pusher') {
            connectToPusher(conn);
        }
    });
}

async function handleCreatePusherConnection() {
    const newConnection = {
        id: `conn_${new Date().getTime()}`,
        type: 'pusher',
        name: document.getElementById("connection-name").value.trim(),
        config: {
            key: document.getElementById("pusher-key").value.trim(),
            secret: document.getElementById("pusher-secret").value.trim(),
            cluster: document.getElementById("pusher-cluster").value.trim(),
            channel: document.getElementById("pusher-channel").value.trim(),
            event: document.getElementById("pusher-event").value.trim()
        }
    };

    if (!newConnection.name || !newConnection.config.key || !newConnection.config.secret || !newConnection.config.cluster || !newConnection.config.channel || !newConnection.config.event) {
        alert("All fields are required to create a connection.");
        return;
    }

    const settings = await getSettings();
    settings.connections.push(newConnection);
    await saveSettings(settings);
    
    hidePusherConfigModal();
    await loadAndRenderConnections(); // Full re-render and connect
}

async function handleConnectionCardAction(event) {
    const button = event.target.closest('button');
    if (!button) return;

    const card = button.closest('.connection-card');
    const connectionId = card.dataset.connectionId;
    
    if (button.dataset.action === 'delete') {
        if (confirm('Are you sure you want to delete this connection?')) {
            const settings = await getSettings();
            settings.connections = settings.connections.filter(c => c.id !== connectionId);
            await saveSettings(settings);

            // Disconnect if active
            const activeInstance = activePusherInstances[connectionId];
            if (activeInstance) {
                activeInstance.pusher.disconnect();
                delete activePusherInstances[connectionId];
            }
            
            await loadAndRenderConnections();
        }
    }
}

function renderConnections(connections) {
    const container = document.getElementById("connections-list-container");
    const noConnectionsMessage = document.getElementById("no-connections-message");
    
    const activeIds = new Set(connections.map(c => c.id));
    // Remove cards for connections that no longer exist
    container.querySelectorAll('.connection-card').forEach(card => {
        if (!activeIds.has(card.dataset.connectionId)) {
            card.remove();
        }
    });

    if (connections.length === 0) {
        noConnectionsMessage.classList.remove('hidden');
    } else {
        noConnectionsMessage.classList.add('hidden');
        connections.forEach(conn => {
            let card = container.querySelector(`.connection-card[data-connection-id="${conn.id}"]`);
            if (!card) {
                card = document.createElement('div');
                card.className = "connection-card bg-white p-4 rounded-lg shadow-sm border border-gray-200";
                card.dataset.connectionId = conn.id;
                container.appendChild(card);
            }
            
            card.innerHTML = `
                <div class="flex justify-between items-start">
                    <div>
                        <div class="flex items-center">
                            <span class="status-dot disconnected"></span>
                            <h3 class="font-bold text-md text-gray-800">${conn.name}</h3>
                        </div>
                        <p class="text-sm text-gray-500 ml-5">${conn.type === 'pusher' ? 'Pusher Real-time' : conn.type}</p>
                    </div>
                    <div>
                        <button data-action="delete" class="p-1 text-gray-400 hover:text-red-600 rounded-full">
                            <svg class="w-5 h-5" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M19 7l-.867 12.142A2 2 0 0116.138 21H7.862a2 2 0 01-1.995-1.858L5 7m5 4v6m4-6v6m1-10V4a1 1 0 00-1-1h-4a1 1 0 00-1 1v3M4 7h16"></path></svg>
                        </button>
                    </div>
                </div>
                <div class="ml-5 mt-2 text-xs text-gray-600" id="status-text-${conn.id}">Status: Not Connected</div>`;
        });
    }
}

// --- PUSHER LOGIC ---
async function createPusherSignature(secret, stringToSign) {
    const encoder = new TextEncoder();
    const keyData = encoder.encode(secret);
    const messageData = encoder.encode(stringToSign);
    const key = await window.crypto.subtle.importKey("raw", keyData, { name: "HMAC", hash: "SHA-256" }, false, ["sign"]);
    const signature = await window.crypto.subtle.sign("HMAC", key, messageData);
    return Array.from(new Uint8Array(signature)).map(b => b.toString(16).padStart(2, '0')).join('');
}

function connectToPusher(connection) {
    const { id, config } = connection;
    
    // Disconnect existing instance for this ID if it exists
    if (activePusherInstances[id]) {
        activePusherInstances[id].pusher.disconnect();
    }

    try {
        updateConnectionStatus(id, "connecting", "Connecting...");
        logToUI(`Initializing Pusher for '${connection.name}'...`);
        
        const pusher = new Pusher(config.key, {
            cluster: config.cluster,
            authorizer: (channel, options) => ({
                authorize: async (socketId, callback) => {
                    const stringToSign = `${socketId}:${channel.name}`;
                    try {
                        const signature = await createPusherSignature(config.secret, stringToSign);
                        const authData = { auth: `${config.key}:${signature}` };
                        callback(null, authData);
                    } catch (error) {
                         logToUI(`Signature failed for '${connection.name}': ${error.message}`, "ERROR");
                         callback(error, null);
                    }
                }
            })
        });

        const channel = pusher.subscribe(config.channel);
        activePusherInstances[id] = { pusher, channel };

        channel.bind('pusher:subscription_succeeded', () => {
            updateConnectionStatus(id, "connected", `Connected & listening on '${config.channel}'`);
            logToUI(`'${connection.name}' subscribed successfully!`, "SUCCESS");
        });

        channel.bind('pusher:subscription_error', (status) => {
            const errorMsg = (status && status.error) ? status.error.message : JSON.stringify(status);
            updateConnectionStatus(id, "error", `Subscription Error: ${errorMsg}`);
            logToUI(`Subscription failed for '${connection.name}': ${errorMsg}`, "ERROR");
        });

        const eventName = config.event.startsWith('client-') ? config.event : `client-${config.event}`;
        channel.bind(eventName, (data) => {
            logToUI(`Event received on '${connection.name}': ${JSON.stringify(data)}`, "SUCCESS");
            if (data && data.name) {
                highlightStudentInSheet(data.name);
            } else {
                logToUI(`Event on '${connection.name}' missing 'name' property.`, "ERROR");
            }
        });

    } catch (error) {
        updateConnectionStatus(id, "error", `Connection Failed: ${error.message}`);
        logToUI(`Pusher error for '${connection.name}': ${error.message}`, "ERROR");
    }
}

function updateConnectionStatus(connectionId, status, message) {
    const card = document.querySelector(`.connection-card[data-connection-id="${connectionId}"]`);
    if (card) {
        const dot = card.querySelector('.status-dot');
        const text = card.querySelector(`#status-text-${connectionId}`);
        
        dot.className = `status-dot ${status}`; // status should be 'connecting', 'connected', 'error', or 'disconnected'
        text.textContent = `Status: ${message}`;
    }
}

// --- EXCEL INTERACTION ---
async function highlightStudentInSheet(studentName) {
    if (!studentName) {
        logToUI("Highlight function called with no student name.", "ERROR");
        return;
    }
    logToUI(`Attempting to highlight student: '${studentName}'...`);
    try {
        await Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getItem(MASTER_LIST_SHEET);
            const usedRange = sheet.getUsedRange();
            
            usedRange.load("values, rowCount, columnCount");
            await context.sync();

            if (!usedRange.values || usedRange.values.length === 0) {
                throw new Error("The 'Master List' sheet is empty or contains no data.");
            }
            
            const headers = usedRange.values[0].map(h => String(h || '').toLowerCase());
            const nameColumnIndex = headers.indexOf("studentname");
            
            if (nameColumnIndex === -1) {
                throw new Error("Could not find a 'StudentName' column in the Master List.");
            }

            const nameColumnRange = usedRange.getColumn(nameColumnIndex);
            logToUI(`Searching for '${studentName}' in the 'StudentName' column.`);

            const searchResult = nameColumnRange.find(studentName, {
                completeMatch: false,
                matchCase: false,
                searchDirection: Excel.SearchDirection.forward
            });
            
            searchResult.load("address, rowIndex");
            await context.sync();

            const foundRowIndex = searchResult.rowIndex;
            logToUI(`Match found for '${studentName}' at row ${foundRowIndex + 1}.`);
            
            const entireRow = sheet.getRangeByIndexes(foundRowIndex, 0, 1, usedRange.columnCount);
            entireRow.format.fill.color = "yellow";
            
            sheet.getCell(foundRowIndex, nameColumnIndex).select();
            
            await context.sync();
            logToUI(`Successfully highlighted row for '${studentName}'.`, "SUCCESS");

        });
    } catch (error) {
        let errorMessage = "An unknown error occurred.";
        if (error instanceof OfficeExtension.Error) {
            errorMessage = error.debugInfo ? (error.debugInfo.message || error.message) : error.message;
            if (error.code === "ItemNotFound") {
                 errorMessage = `Could not find a student matching '${studentName}' in the 'StudentName' column.`;
            } else if (error.code === "WorksheetNotFound") {
                errorMessage = `The '${MASTER_LIST_SHEET}' worksheet could not be found.`;
            }
        } else {
            errorMessage = error.message;
        }
        logToUI(`Error during highlight: ${errorMessage}`, "ERROR");
        console.error("Error highlighting student:", error);
    }
}

