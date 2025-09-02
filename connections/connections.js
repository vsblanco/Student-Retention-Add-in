'use strict';

// --- STATE & CONSTANTS ---
const SETTINGS_KEY = "studentRetentionSettings";
const MASTER_LIST_SHEET = "Master List";
const LOGO_MAP = {
    pusher: '../images/pusher-icon.png',
    'power-automate': '../images/power-automate-icon.png'
};

let activePusherInstances = {};
let connectionToModifyId = null; 

// --- INITIALIZATION ---
Office.onReady((info) => {
    if (info.host === Office.HostType.Excel) {
        initializeEventListeners();
        logToUI("Connections pane ready.");
        loadAndRenderConnections();
        // Close dropdowns if user clicks outside
        document.addEventListener('click', (event) => {
            if (!event.target.closest('.options-menu-container')) {
                document.querySelectorAll('.options-menu').forEach(menu => menu.classList.remove('open'));
            }
        });
    }
});

function initializeEventListeners() {
    // Main UI
    document.getElementById("new-connection-button").onclick = showSelectServiceModal;
    
    // Select Service Modal
    document.getElementById("cancel-select-service-button").onclick = hideSelectServiceModal;
    document.getElementById("select-pusher-button").onclick = showPusherConfigModal;
    
    // Pusher Config Modal (Create)
    document.getElementById("cancel-pusher-config-button").onclick = hidePusherConfigModal;
    document.getElementById("create-pusher-connection-button").onclick = handleCreatePusherConnection;
    
    // Pusher Config Modal (Edit)
    document.getElementById("cancel-edit-pusher-config-button").onclick = hideEditPusherConfigModal;
    document.getElementById("update-pusher-connection-button").onclick = handleUpdateConnection;

    // Add Action Modal
    document.getElementById("add-action-modal").addEventListener('click', handleAddActionClick);
    document.getElementById("cancel-add-action-button").onclick = hideAddActionModal;

    // Confirm Delete Modal
    document.getElementById("cancel-delete-button").onclick = hideDeleteConfirmModal;
    document.getElementById("confirm-delete-button").onclick = handleConfirmDelete;

    // Log panel
    document.getElementById("log-header").onclick = toggleLogPanel;
    document.getElementById("clear-log-button").onclick = clearLogPanel;

    // Event delegation for connection cards
    document.getElementById("connections-list-container").addEventListener('click', handleConnectionCardAction);
}

// --- LOGGING ---
function logToUI(message, type = 'INFO') {
    const logContainer = document.getElementById('log-container');
    if (!logContainer) return;
    const p = document.createElement('p');
    const timestamp = new Date().toLocaleTimeString();
    p.innerHTML = `<span class="text-gray-500">${timestamp} [${type}]:</span> ${message}`;
    if (type === 'ERROR') p.className = 'text-red-400';
    else if (type === 'SUCCESS') p.className = 'text-green-400';
    else p.className = 'text-gray-300';
    logContainer.appendChild(p);
    logContainer.scrollTop = logContainer.scrollHeight;
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
function showEditPusherConfigModal() { document.getElementById("edit-pusher-config-modal").classList.remove('hidden'); }
function hideEditPusherConfigModal() { document.getElementById("edit-pusher-config-modal").classList.add('hidden'); connectionToModifyId = null; }
function showAddActionModal(connectionId) { connectionToModifyId = connectionId; document.getElementById("add-action-modal").classList.remove('hidden'); }
function hideAddActionModal() { document.getElementById("add-action-modal").classList.add('hidden'); connectionToModifyId = null; }
function showDeleteConfirmModal(connectionId) { connectionToModifyId = connectionId; document.getElementById("confirm-delete-modal").classList.remove('hidden'); }
function hideDeleteConfirmModal() { document.getElementById("confirm-delete-modal").classList.add('hidden'); connectionToModifyId = null; }


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
            if (settings.connections && Array.isArray(settings.connections)) {
                settings.connections.forEach(conn => { conn.actions = conn.actions || []; });
            }
            return { ...defaults, ...settings };
        } catch (e) { logToUI("Error parsing settings.", "ERROR"); return defaults; }
    }
    return defaults;
}

async function saveSettings(settingsToSave) {
    Office.context.document.settings.set(SETTINGS_KEY, JSON.stringify(settingsToSave));
    await new Promise((resolve, reject) => {
        Office.context.document.settings.saveAsync(result => {
            if (result.status === Office.AsyncResultStatus.Failed) reject(new Error("Failed to save settings: " + result.error.message));
            else { logToUI("Settings saved successfully."); resolve(); }
        });
    });
}

// --- CONNECTION & ACTION MANAGEMENT ---
async function loadAndRenderConnections() {
    logToUI("Loading all connections...");
    const settings = await getSettings();
    renderConnections(settings.connections);
    settings.connections.forEach(conn => {
        if (conn.type === 'pusher') connectToPusher(conn);
    });
}

async function handleCreatePusherConnection() {
    const newConnection = {
        id: `conn_${new Date().getTime()}`, type: 'pusher',
        name: document.getElementById("connection-name").value.trim(),
        config: {
            key: document.getElementById("pusher-key").value.trim(),
            secret: document.getElementById("pusher-secret").value.trim(),
            cluster: document.getElementById("pusher-cluster").value.trim(),
            channel: document.getElementById("pusher-channel").value.trim(),
            event: document.getElementById("pusher-event").value.trim()
        },
        actions: []
    };
    if (!newConnection.name || !newConnection.config.key || !newConnection.config.secret || !newConnection.config.cluster || !newConnection.config.channel || !newConnection.config.event) {
        alert("All fields are required."); return;
    }
    const settings = await getSettings();
    settings.connections.push(newConnection);
    await saveSettings(settings);
    hidePusherConfigModal();
    await loadAndRenderConnections();
}

async function handleUpdateConnection() {
    const settings = await getSettings();
    const conn = settings.connections.find(c => c.id === connectionToModifyId);
    if (!conn) return;

    conn.name = document.getElementById("edit-connection-name").value.trim();
    conn.config.key = document.getElementById("edit-pusher-key").value.trim();
    conn.config.cluster = document.getElementById("edit-pusher-cluster").value.trim();
    conn.config.channel = document.getElementById("edit-pusher-channel").value.trim();
    conn.config.event = document.getElementById("edit-pusher-event").value.trim();
    
    const newSecret = document.getElementById("edit-pusher-secret").value.trim();
    if (newSecret) conn.config.secret = newSecret;
    
    await saveSettings(settings);
    hideEditPusherConfigModal();
    await loadAndRenderConnections();
}

async function handleConfirmDelete() {
    const settings = await getSettings();
    settings.connections = settings.connections.filter(c => c.id !== connectionToModifyId);
    await saveSettings(settings);
    const activeInstance = activePusherInstances[connectionToModifyId];
    if (activeInstance) {
        activeInstance.pusher.disconnect();
        delete activePusherInstances[connectionToModifyId];
    }
    hideDeleteConfirmModal();
    await loadAndRenderConnections();
}


async function handleConnectionCardAction(event) {
    const button = event.target.closest('button[data-action]');
    if (!button) return;

    event.stopPropagation();
    const card = button.closest('.connection-card');
    const connectionId = card.dataset.connectionId;
    const action = button.dataset.action;

    if (action === 'toggle-menu') {
        const menu = button.nextElementSibling;
        const isMenuOpen = menu.classList.contains('open');
        document.querySelectorAll('.options-menu').forEach(m => m.classList.remove('open'));
        if (!isMenuOpen) menu.classList.add('open');
    } else if (action === 'show-edit-modal') {
        const settings = await getSettings();
        const conn = settings.connections.find(c => c.id === connectionId);
        if (conn) {
            connectionToModifyId = connectionId;
            document.getElementById('edit-connection-name').value = conn.name;
            document.getElementById('edit-pusher-key').value = conn.config.key;
            document.getElementById('edit-pusher-secret').value = ''; // Don't pre-fill secret
            document.getElementById('edit-pusher-cluster').value = conn.config.cluster;
            document.getElementById('edit-pusher-channel').value = conn.config.channel;
            document.getElementById('edit-pusher-event').value = conn.config.event;
            showEditPusherConfigModal();
        }
    } else if (action === 'show-delete-modal') {
        showDeleteConfirmModal(connectionId);
    } else if (action === 'add-action') {
        showAddActionModal(connectionId);
    } else if (action === 'delete-action') {
        const actionId = button.dataset.actionId;
        const settings = await getSettings();
        const conn = settings.connections.find(c => c.id === connectionId);
        if (conn) {
            conn.actions = conn.actions.filter(a => a.id !== actionId);
            await saveSettings(settings);
            await loadAndRenderConnections();
        }
    }
}

async function handleAddActionClick(event) {
    const button = event.target.closest('button[data-action-type]');
    if (!button) return;

    const actionType = button.dataset.actionType;
    if (actionType === 'liveHighlight') {
        const newAction = { id: `action_${new Date().getTime()}`, type: 'liveHighlight', name: 'Live Submission Highlighting' };
        const settings = await getSettings();
        const connection = settings.connections.find(c => c.id === connectionToModifyId);
        if (connection) {
            if (connection.actions.some(a => a.type === 'liveHighlight')) {
                alert('This connection already has this action.'); return;
            }
            connection.actions.push(newAction);
            await saveSettings(settings);
            hideAddActionModal();
            await loadAndRenderConnections();
        }
    }
}

function renderConnections(connections) {
    const container = document.getElementById("connections-list-container");
    const noConnectionsMessage = document.getElementById("no-connections-message");
    container.innerHTML = '';
    container.appendChild(noConnectionsMessage);

    if (!connections || connections.length === 0) {
        noConnectionsMessage.classList.remove('hidden');
    } else {
        noConnectionsMessage.classList.add('hidden');
        connections.forEach(conn => {
            const card = document.createElement('div');
            card.className = "connection-card bg-white p-4 rounded-lg shadow-sm border border-gray-200";
            card.dataset.connectionId = conn.id;

            let actionsHtml = '<p class="text-xs text-gray-400 italic">No actions configured.</p>';
            if (conn.actions && conn.actions.length > 0) {
                actionsHtml = conn.actions.map(action => `
                    <div class="flex items-center justify-between text-sm text-gray-700 py-1 pl-2 border-l-2 border-blue-200">
                        <span>${action.name}</span>
                        <button data-action="delete-action" data-action-id="${action.id}" class="p-1 text-gray-400 hover:text-red-600 rounded-full">
                            <svg class="w-4 h-4" fill="currentColor" viewBox="0 0 20 20"><path fill-rule="evenodd" d="M4.293 4.293a1 1 0 011.414 0L10 8.586l4.293-4.293a1 1 0 111.414 1.414L11.414 10l4.293 4.293a1 1 0 01-1.414 1.414L10 11.414l-4.293 4.293a1 1 0 01-1.414-1.414L8.586 10 4.293 5.707a1 1 0 010-1.414z" clip-rule="evenodd"></path></svg>
                        </button>
                    </div>
                `).join('');
            }
            
            const logoSrc = LOGO_MAP[conn.type] || '';
            const logoHtml = logoSrc ? `<img src="${logoSrc}" alt="${conn.type} logo" class="h-6 w-6 mr-2">` : '';

            card.innerHTML = `
                <div class="flex justify-between items-center">
                    <div class="flex items-center">
                         <div class="relative has-tooltip">
                            <span id="status-dot-${conn.id}" class="status-dot disconnected"></span>
                            <div id="status-tooltip-${conn.id}" class="tooltip absolute bottom-full mb-2 w-max px-2 py-1 bg-gray-700 text-white text-xs rounded">Status: Not Connected</div>
                        </div>
                        ${logoHtml}
                        <h3 class="font-bold text-md text-gray-800">${conn.name}</h3>
                    </div>
                    <div class="relative options-menu-container">
                        <button data-action="toggle-menu" class="p-2 text-gray-500 hover:bg-gray-200 rounded-full">
                            <svg class="w-5 h-5" fill="currentColor" viewBox="0 0 20 20"><path d="M10 6a2 2 0 110-4 2 2 0 010 4zM10 12a2 2 0 110-4 2 2 0 010 4zM10 18a2 2 0 110-4 2 2 0 010 4z"></path></svg>
                        </button>
                        <div class="options-menu absolute right-0 mt-2 w-32 bg-white rounded-md shadow-lg border z-10">
                            <button data-action="show-edit-modal" class="block w-full text-left px-4 py-2 text-sm text-gray-700 hover:bg-gray-100">Edit</button>
                            <button data-action="show-delete-modal" class="block w-full text-left px-4 py-2 text-sm text-red-600 hover:bg-gray-100">Delete</button>
                        </div>
                    </div>
                </div>
                <div class="mt-4 ml-5 pl-4 border-l-2 border-gray-200">
                    <h4 class="text-xs font-semibold text-gray-500 uppercase mb-2">Configured Actions</h4>
                    ${actionsHtml}
                    <button data-action="add-action" class="mt-2 text-xs text-blue-600 hover:underline">+ Add Action</button>
                </div>`;
            container.appendChild(card);
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
    const { id, config, actions } = connection;
    if (activePusherInstances[id]) activePusherInstances[id].pusher.disconnect();
    if (!actions || actions.length === 0) {
        logToUI(`Connection '${connection.name}' has no actions configured. Skipping.`);
        updateConnectionStatus(id, 'disconnected', 'Ready. Add an action to begin listening.');
        return;
    }
    if (!config.channel || !config.event) {
        logToUI(`Connection '${connection.name}' is missing channel or event name.`, "ERROR");
        updateConnectionStatus(id, 'error', 'Missing Channel or Event name');
        return;
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
                        callback(null, { auth: `${config.key}:${signature}` });
                    } catch (error) {
                        logToUI(`Signature failed for '${connection.name}': ${error.message}`, "ERROR");
                        callback(error, null);
                    }
                }
            })
        });
        const channel = pusher.subscribe(config.channel);
        channel.bind('pusher:subscription_succeeded', () => {
            updateConnectionStatus(id, "connected", `Connected & listening on '${config.channel}'`);
            logToUI(`'${connection.name}' subscribed successfully!`, "SUCCESS");
        });
        channel.bind('pusher:subscription_error', (status) => {
            const errorMsg = status?.error?.message || JSON.stringify(status);
            updateConnectionStatus(id, "error", `Subscription Error: ${errorMsg}`);
            logToUI(`Subscription failed for '${connection.name}': ${errorMsg}`, "ERROR");
        });
        const eventName = config.event.startsWith('client-') ? config.event : `client-${config.event}`;
        channel.bind(eventName, (data) => {
            logToUI(`Event received on '${connection.name}': ${JSON.stringify(data)}`, "SUCCESS");
            if (actions.some(a => a.type === 'liveHighlight') && data && data.name) {
                highlightStudentInSheet(data.name);
            }
        });
        activePusherInstances[id] = { pusher, channel };
    } catch (error) {
        updateConnectionStatus(id, "error", `Connection Failed: ${error.message}`);
        logToUI(`Pusher error for '${connection.name}': ${error.message}`, "ERROR");
    }
}

function updateConnectionStatus(connectionId, status, message) {
    const dot = document.getElementById(`status-dot-${connectionId}`);
    const tooltip = document.getElementById(`status-tooltip-${connectionId}`);
    if (dot && tooltip) {
        dot.className = `status-dot ${status}`;
        tooltip.textContent = `Status: ${message}`;
    }
}

// --- EXCEL INTERACTION ---
async function highlightStudentInSheet(studentName) {
    logToUI(`Attempting to highlight student: '${studentName}'...`);
    try {
        await Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getItem(MASTER_LIST_SHEET);
            const usedRange = sheet.getUsedRange();
            usedRange.load("values, rowCount, columnCount");
            await context.sync();
            if (!usedRange.values || usedRange.values.length === 0) throw new Error("'Master List' sheet is empty.");
            const headers = usedRange.values[0].map(h => String(h || '').toLowerCase());
            const nameColumnIndex = headers.indexOf("studentname");
            if (nameColumnIndex === -1) throw new Error("Could not find a 'StudentName' column.");
            const searchResult = usedRange.getColumn(nameColumnIndex).find(studentName, { completeMatch: false, matchCase: false, searchDirection: Excel.SearchDirection.forward });
            searchResult.load("rowIndex");
            await context.sync();
            const foundRowIndex = searchResult.rowIndex;
            logToUI(`Match found for '${studentName}' at row ${foundRowIndex + 1}.`);
            sheet.getRangeByIndexes(foundRowIndex, 0, 1, usedRange.columnCount).format.fill.color = "yellow";
            sheet.getCell(foundRowIndex, nameColumnIndex).select();
            await context.sync();
            logToUI(`Successfully highlighted row for '${studentName}'.`, "SUCCESS");
        });
    } catch (error) {
        let errorMessage = error.message;
        if (error instanceof OfficeExtension.Error) {
            errorMessage = error.debugInfo?.message || error.message;
            if (error.code === "ItemNotFound") errorMessage = `Could not find a student matching '${studentName}'.`;
            else if (error.code === "WorksheetNotFound") errorMessage = `The '${MASTER_LIST_SHEET}' worksheet could not be found.`;
        }
        logToUI(`Error during highlight: ${errorMessage}`, "ERROR");
    }
}

