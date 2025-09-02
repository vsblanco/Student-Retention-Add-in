'use strict';

// --- STATE & CONSTANTS ---
const SETTINGS_KEY = "studentRetentionSettings";
const LOGO_MAP = {
    pusher: '../images/pusher-icon.png',
    'power-automate': '../images/power-automate-icon.png'
};

let activePusherInstances = {};
let connectionToModifyId = null; 
let actionToModifyId = null;

// --- INITIALIZATION ---
Office.onReady((info) => {
    if (info.host === Office.HostType.Excel) {
        initializeEventListeners();
        logToUI("Connections pane ready.");
        loadAndRenderConnections();
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
    
    // Pusher Config Modals
    document.getElementById("cancel-pusher-config-button").onclick = hidePusherConfigModal;
    document.getElementById("create-pusher-connection-button").onclick = handleCreatePusherConnection;
    document.getElementById("cancel-edit-pusher-config-button").onclick = hideEditPusherConfigModal;
    document.getElementById("update-pusher-connection-button").onclick = handleUpdateConnection;

    // Add Action Modal
    document.getElementById("add-action-modal").addEventListener('click', handleAddActionClick);
    document.getElementById("cancel-add-action-button").onclick = hideAddActionModal;

    // Action Settings Modal
    document.getElementById("cancel-action-settings-button").onclick = hideActionSettingsModal;
    document.getElementById("save-action-settings-button").onclick = handleSaveActionSettings;
    document.getElementById("color-picker").addEventListener('click', handleColorPicker);
    document.getElementById("sheet-options").addEventListener('change', handleSheetOptions);


    // Confirm Delete Modal
    document.getElementById("cancel-delete-button").onclick = hideDeleteConfirmModal;
    document.getElementById("confirm-delete-button").onclick = handleConfirmDelete;

    // Log panel
    document.getElementById("log-header").onclick = toggleLogPanel;
    document.getElementById("clear-log-button").onclick = clearLogPanel;

    // Event delegation for cards
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

function clearLogPanel(e) { if (e) e.stopPropagation(); document.getElementById('log-container').innerHTML = ''; logToUI("Log cleared."); }

// --- MODAL & UI MANAGEMENT ---
function showSelectServiceModal() { document.getElementById("select-service-modal").classList.remove('hidden'); }
function hideSelectServiceModal() { document.getElementById("select-service-modal").classList.add('hidden'); }
function showPusherConfigModal() { hideSelectServiceModal(); document.getElementById("pusher-config-modal").classList.remove('hidden'); }
function hidePusherConfigModal() { document.getElementById("pusher-config-modal").classList.add('hidden'); resetPusherForm(); }
function showEditPusherConfigModal() { document.getElementById("edit-pusher-config-modal").classList.remove('hidden'); }
function hideEditPusherConfigModal() { document.getElementById("edit-pusher-config-modal").classList.add('hidden'); connectionToModifyId = null; }
function showAddActionModal(connectionId) { connectionToModifyId = connectionId; document.getElementById("add-action-modal").classList.remove('hidden'); }
function hideAddActionModal() { document.getElementById("add-action-modal").classList.add('hidden'); connectionToModifyId = null; }
function showActionSettingsModal() { document.getElementById("highlight-action-settings-modal").classList.remove('hidden'); }
function hideActionSettingsModal() { 
    document.getElementById("highlight-action-settings-modal").classList.add('hidden');
    // Also hide the underlying add action modal
    document.getElementById("add-action-modal").classList.add('hidden');
    connectionToModifyId = null; 
    actionToModifyId = null; 
}
function showDeleteConfirmModal(connectionId) { connectionToModifyId = connectionId; document.getElementById("confirm-delete-modal").classList.remove('hidden'); }
function hideDeleteConfirmModal() { document.getElementById("confirm-delete-modal").classList.add('hidden'); connectionToModifyId = null; }

function resetPusherForm() { ["connection-name", "pusher-key", "pusher-secret", "pusher-cluster", "pusher-channel", "pusher-event"].forEach(id => document.getElementById(id).value = ''); }
function getTodaysLdaSheetName() { const now = new Date(); return `LDA ${now.getMonth() + 1}-${now.getDate()}-${now.getFullYear()}`; }

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
    settings.connections.forEach(conn => { if (conn.type === 'pusher') connectToPusher(conn); });
}

async function handleCreatePusherConnection() {
    const newConnection = {
        id: `conn_${new Date().getTime()}`, type: 'pusher',
        name: document.getElementById("connection-name").value.trim(),
        config: {
            key: document.getElementById("pusher-key").value.trim(), secret: document.getElementById("pusher-secret").value.trim(),
            cluster: document.getElementById("pusher-cluster").value.trim(), channel: document.getElementById("pusher-channel").value.trim(),
            event: document.getElementById("pusher-event").value.trim()
        }, actions: []
    };
    if (Object.values(newConnection.config).some(v => !v) || !newConnection.name) {
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
    if (activeInstance) { activeInstance.pusher.disconnect(); delete activePusherInstances[connectionToModifyId]; }
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
            document.getElementById('edit-pusher-secret').value = '';
            document.getElementById('edit-pusher-cluster').value = conn.config.cluster;
            document.getElementById('edit-pusher-channel').value = conn.config.channel;
            document.getElementById('edit-pusher-event').value = conn.config.event;
            showEditPusherConfigModal();
        }
    } else if (action === 'show-delete-modal') {
        showDeleteConfirmModal(connectionId);
    } else if (action === 'add-action') {
        showAddActionModal(connectionId);
    } else if (action === 'edit-action') {
        const actionId = button.dataset.actionId;
        const settings = await getSettings();
        const conn = settings.connections.find(c => c.id === connectionId);
        const actionToEdit = conn?.actions.find(a => a.id === actionId);
        if (actionToEdit) {
            connectionToModifyId = connectionId;
            actionToModifyId = actionToEdit.id;
            populateAndShowActionSettingsModal(actionToEdit.config);
        }
    } else if (action === 'delete-action') {
        const actionId = button.dataset.actionId;
        const settings = await getSettings();
        const conn = settings.connections.find(c => c.id === connectionId);
        if (conn) { conn.actions = conn.actions.filter(a => a.id !== actionId); await saveSettings(settings); await loadAndRenderConnections(); }
    }
}

async function handleAddActionClick(event) {
    const button = event.target.closest('button[data-action-type]');
    if (!button) return;
    const actionType = button.dataset.actionType;
    if (actionType === 'liveHighlight') {
        const settings = await getSettings();
        const connection = settings.connections.find(c => c.id === connectionToModifyId);
        if (connection && connection.actions.some(a => a.type === 'liveHighlight')) {
            alert('This connection already has this action.'); return;
        }
        populateAndShowActionSettingsModal();
    }
}

function populateAndShowActionSettingsModal(config) {
    const defaults = { color: '#92d050', sheetType: 'today', customSheetName: '', ignoreColumns: '' };
    const settings = { ...defaults, ...config };
    document.querySelectorAll('.color-swatch').forEach(swatch => {
        swatch.classList.toggle('selected', swatch.dataset.color === settings.color);
    });
    document.getElementById('custom-color-input').value = settings.color;
    document.getElementById(`sheet-${settings.sheetType}`).checked = true;
    const customSheetInput = document.getElementById('custom-sheet-name');
    customSheetInput.classList.toggle('hidden', settings.sheetType !== 'custom');
    customSheetInput.value = settings.customSheetName;
    document.getElementById('ignore-columns').value = settings.ignoreColumns;
    showActionSettingsModal();
}

async function handleSaveActionSettings() {
    const settings = await getSettings();
    const connection = settings.connections.find(c => c.id === connectionToModifyId);
    if (!connection) return;

    const color = document.getElementById('custom-color-input').value.trim() || '#92d050';
    const sheetType = document.querySelector('input[name="sheet-choice"]:checked').value;
    const customSheetName = document.getElementById('custom-sheet-name').value.trim();
    const ignoreColumns = document.getElementById('ignore-columns').value.trim();
    const actionConfig = { color, sheetType, customSheetName, ignoreColumns };

    if (actionToModifyId) {
        const action = connection.actions.find(a => a.id === actionToModifyId);
        if (action) action.config = actionConfig;
    } else {
        const newAction = { id: `action_${new Date().getTime()}`, type: 'liveHighlight', name: 'Auto Highlight', config: actionConfig };
        connection.actions.push(newAction);
    }
    await saveSettings(settings);
    hideActionSettingsModal();
    await loadAndRenderConnections();
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
                    <div class="flex items-center justify-between text-sm text-gray-700 py-1 pl-2 border-l-2 border-blue-200 group">
                        <span>${action.name}</span>
                        <div class="flex items-center opacity-0 group-hover:opacity-100 transition-opacity">
                            <button data-action="edit-action" data-action-id="${action.id}" class="p-1 text-gray-400 hover:text-blue-600 rounded-full"><svg class="w-4 h-4" fill="currentColor" viewBox="0 0 20 20"><path d="M13.586 3.586a2 2 0 112.828 2.828l-.793.793-2.828-2.828.793-.793zM11.379 5.793L3 14.172V17h2.828l8.38-8.379-2.83-2.828z"></path></svg></button>
                            <button data-action="delete-action" data-action-id="${action.id}" class="p-1 text-gray-400 hover:text-red-600 rounded-full"><svg class="w-4 h-4" fill="currentColor" viewBox="0 0 20 20"><path fill-rule="evenodd" d="M4.293 4.293a1 1 0 011.414 0L10 8.586l4.293-4.293a1 1 0 111.414 1.414L11.414 10l4.293 4.293a1 1 0 01-1.414 1.414L10 11.414l-4.293 4.293a1 1 0 01-1.414-1.414L8.586 10 4.293 5.707a1 1 0 010-1.414z" clip-rule="evenodd"></path></svg></button>
                        </div>
                    </div>`).join('');
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
                        <button data-action="toggle-menu" class="p-2 text-gray-500 hover:bg-gray-200 rounded-full"><svg class="w-5 h-5" fill="currentColor" viewBox="0 0 20 20"><path d="M10 6a2 2 0 110-4 2 2 0 010 4zM10 12a2 2 0 110-4 2 2 0 010 4zM10 18a2 2 0 110-4 2 2 0 010 4z"></path></svg></button>
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

// --- Action Settings UI Handlers ---
function handleColorPicker(event) {
    if (event.target.classList.contains('color-swatch')) {
        document.querySelectorAll('.color-swatch').forEach(swatch => swatch.classList.remove('selected'));
        event.target.classList.add('selected');
        document.getElementById('custom-color-input').value = event.target.dataset.color;
    }
}
function handleSheetOptions(event) {
    const customSheetInput = document.getElementById('custom-sheet-name');
    customSheetInput.classList.toggle('hidden', event.target.value !== 'custom');
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
                    try {
                        const signature = await createPusherSignature(config.secret, `${socketId}:${channel.name}`);
                        callback(null, { auth: `${config.key}:${signature}` });
                    } catch (error) { callback(error, null); }
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
        });
        const eventName = config.event.startsWith('client-') ? config.event : `client-${config.event}`;
        channel.bind(eventName, (data) => {
            logToUI(`Event received on '${connection.name}': ${JSON.stringify(data)}`, "SUCCESS");
            const highlightAction = actions.find(a => a.type === 'liveHighlight');
            if (highlightAction && data && data.name) {
                highlightStudentInSheet(data.name, highlightAction.config);
            }
        });
        activePusherInstances[id] = { pusher, channel };
    } catch (error) { updateConnectionStatus(id, "error", `Connection Failed: ${error.message}`); }
}

function updateConnectionStatus(connectionId, status, message) {
    const dot = document.getElementById(`status-dot-${connectionId}`);
    const tooltip = document.getElementById(`status-tooltip-${connectionId}`);
    if (dot && tooltip) { dot.className = `status-dot ${status}`; tooltip.textContent = `Status: ${message}`; }
}

// --- EXCEL INTERACTION ---
function reformatName(name) {
    const parts = name.split(',').map(p => p.trim());
    if (parts.length === 2) {
        return `${parts[1]} ${parts[0]}`;
    } else {
        const spaceIndex = name.lastIndexOf(' ');
        if (spaceIndex !== -1) {
            const first = name.substring(0, spaceIndex);
            const last = name.substring(spaceIndex + 1);
            return `${last}, ${first}`;
        }
    }
    return name;
}

async function highlightStudentInSheet(studentName, actionConfig) {
    const config = { color: '#92d050', sheetType: 'today', ignoreColumns: '', ...actionConfig };
    let sheetName = "Master List";
    if (config.sheetType === 'today') sheetName = getTodaysLdaSheetName();
    else if (config.sheetType === 'custom' && config.customSheetName) sheetName = config.customSheetName;

    logToUI(`Highlighting '${studentName}' on sheet '${sheetName}' with color ${config.color}...`);
    try {
        await Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getItem(sheetName);

            const autoFilter = sheet.getAutoFilter();
            autoFilter.load("isNullObject");
            await context.sync();

            if (!autoFilter.isNullObject) {
                logToUI(`Active filter found on sheet '${sheetName}'. Clearing filter to ensure visibility.`);
                autoFilter.clearCriteria();
            }

            const usedRange = sheet.getUsedRange();
            usedRange.load("values, rowCount, columnCount");
            await context.sync();

            if (!usedRange.values || usedRange.values.length === 0) throw new Error(`Sheet '${sheetName}' is empty.`);
            
            const headers = usedRange.values[0].map(h => String(h || '').toLowerCase());
            const nameColumnIndex = headers.indexOf("studentname");
            if (nameColumnIndex === -1) throw new Error(`Could not find a 'StudentName' column in '${sheetName}'.`);
            
            const nameColumn = usedRange.getColumn(nameColumnIndex);
            let searchResult;

            try {
                searchResult = nameColumn.find(studentName, { completeMatch: false, matchCase: false, searchDirection: Excel.SearchDirection.forward });
                searchResult.load("rowIndex");
                await context.sync();
            } catch (error) {
                if (error instanceof OfficeExtension.Error && error.code === "ItemNotFound") {
                    const alternateName = reformatName(studentName);
                    logToUI(`'${studentName}' not found. Trying alternate format: '${alternateName}'...`);
                    searchResult = nameColumn.find(alternateName, { completeMatch: false, matchCase: false, searchDirection: Excel.SearchDirection.forward });
                    searchResult.load("rowIndex");
                    await context.sync();
                } else {
                    throw error;
                }
            }

            const foundRowIndex = searchResult.rowIndex;
            const ignoredHeaders = new Set((config.ignoreColumns || '').toLowerCase().split(',').map(h => h.trim()).filter(Boolean));
            const ignoredColumnIndices = new Set(headers.map((h, i) => ignoredHeaders.has(h) ? i : -1).filter(i => i !== -1));
            let startCol = -1;
            for (let i = 0; i < usedRange.columnCount; i++) {
                if (!ignoredColumnIndices.has(i) && startCol === -1) {
                    startCol = i;
                } else if (ignoredColumnIndices.has(i) && startCol !== -1) {
                    sheet.getRangeByIndexes(foundRowIndex, startCol, 1, i - startCol).format.fill.color = config.color;
                    startCol = -1;
                }
            }
            if (startCol !== -1) {
                sheet.getRangeByIndexes(foundRowIndex, startCol, 1, usedRange.columnCount - startCol).format.fill.color = config.color;
            }
            sheet.getCell(foundRowIndex, nameColumnIndex).select();
            await context.sync();
            logToUI(`Successfully highlighted row for '${studentName}'.`, "SUCCESS");
        });
    } catch (error) {
        let errorMessage = error.message;
        if (error instanceof OfficeExtension.Error) {
            errorMessage = error.debugInfo?.message || error.message;
            if (error.code === "ItemNotFound") errorMessage = `Could not find student matching '${studentName}' in sheet '${sheetName}'.`;
            else if (error.code === "WorksheetNotFound") errorMessage = `The worksheet '${sheetName}' could not be found.`;
        }
        logToUI(`Error during highlight: ${errorMessage}`, "ERROR");
    }
}

