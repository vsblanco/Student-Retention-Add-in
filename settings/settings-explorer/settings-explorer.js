// Timestamp: 2025-10-02 12:31 PM | Version: 1.4.4
Office.onReady((info) => {
    console.log("Office is ready. Host:", info.host);
    if (info.host === Office.HostType.Excel) {
        initializeSettingsExplorer();
    }
});

/**
 * Initializes the settings explorer page, handling the rendering of the 
 * settings data into a modern tree view and the close button functionality.
 */
function initializeSettingsExplorer() {
    console.log("Initializing Settings Explorer...");
    // --- Get DOM Elements ---
    const closeButton = document.getElementById('close-settings-explorer-button');
    const treeContainer = document.getElementById('settings-tree-container');

    // --- SVG Icons for a richer UI ---
    const ICONS = {
        chevron: `<svg xmlns="http://www.w.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round" class="tree-icon chevron"><polyline points="9 18 15 12 9 6"></polyline></svg>`,
        object: `<svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" class="tree-icon object"><path d="M22 19a2 2 0 0 1-2 2H4a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h5l2 3h9a2 2 0 0 1 2 2z"></path></svg>`,
        array: `<svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" class="tree-icon array"><rect x="3" y="3" width="18" height="18" rx="2" ry="2"></rect><line x1="8" y1="3" x2="8" y2="21"></line><line x1="16" y1="3" x2="16" y2="21"></line></svg>`,
        string: `<svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" class="tree-icon string"><path d="M4 12h16M4 18h16M4 6h16"/></svg>`,
        number: `<svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" class="tree-icon number"><line x1="5" y1="12" x2="19" y2="12"></line><line x1="12" y1="5" x2="12" y2="19"></line></svg>`,
        boolean: `<svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" class="tree-icon boolean"><polyline points="20 6 9 17 4 12"></polyline></svg>`,
        null: `<svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" class="tree-icon null"><circle cx="12" cy="12" r="10"></circle><line x1="4.93" y1="4.93" x2="19.07" y2="19.07"></line></svg>`,
        copy: `<svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect x="9" y="9" width="13" height="13" rx="2" ry="2"></rect><path d="M5 15H4a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2h9a2 2 0 0 1 2 2v1"></path></svg>`,
        check: `<svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="3" stroke-linecap="round" stroke-linejoin="round"><polyline points="20 6 9 17 4 12"></polyline></svg>`
    };

    /**
     * Fetches all settings from the workbook and renders them in the tree view.
     */
    const populateTree = async () => {
        console.log("Attempting to populate the settings tree...");
        treeContainer.innerHTML = ''; // Clear previous content
        
        try {
            await Excel.run(async (context) => {
                console.log("Excel.run() context started.");
                const allSettings = {};
                let hasSettings = false;
                
                const workbookSettings = context.workbook.settings;
                const settingItems = {};

                // 1. Get all setting items and queue up a 'load' command for their values.
                console.log("Loading settings keys from CONSTANTS:", CONSTANTS.SETTINGS_KEYS);
                const settingKeysToLoad = Object.values(CONSTANTS.SETTINGS_KEYS);
                console.log(`Found ${settingKeysToLoad.length} setting keys to load.`);

                for (const settingKey of settingKeysToLoad) {
                    console.log(`- Queuing load for setting key: "${settingKey}"`);
                    const settingItem = workbookSettings.getItemOrNullObject(settingKey);
                    settingItem.load("value");
                    settingItems[settingKey] = settingItem;
                }
                
                // 2. Sync to execute all the queued 'load' commands in one network call.
                console.log("Executing context.sync() to fetch settings values...");
                await context.sync();
                console.log("context.sync() completed successfully.");

                // 3. Now that values are loaded, process them.
                console.log("Processing loaded settings...");
                for (const settingKey in settingItems) {
                    if (Object.prototype.hasOwnProperty.call(settingItems, settingKey)) {
                        const settingItem = settingItems[settingKey];
                        if (settingItem.value) { 
                            console.log(`Found value for setting key "${settingKey}".`);
                            const settingsString = settingItem.value;
                            try {
                                allSettings[settingKey] = JSON.parse(settingsString);
                                hasSettings = true;
                                console.log(`  -> Successfully parsed JSON for "${settingKey}".`);
                            } catch (e) {
                                console.warn(`Could not parse setting for key "${settingKey}":`, e);
                                allSettings[settingKey] = "[Error: Invalid JSON]";
                                hasSettings = true; // Still show the error in the UI
                            }
                        } else {
                            console.log(`No value found for setting key "${settingKey}". It may not be set yet.`);
                        }
                    }
                }

                console.log("Final compiled settings object:", allSettings);

                if (hasSettings) {
                    console.log("Settings found. Building tree view...");
                    const rootUl = document.createElement('ul');
                    rootUl.className = 'settings-tree';
                    buildTree(allSettings, rootUl);
                    treeContainer.appendChild(rootUl);
                    console.log("Tree view built and appended to the DOM.");
                } else {
                    console.log("No settings were found in the workbook.");
                    treeContainer.innerHTML = `<div class="explorer-empty">No settings have been saved for this add-in yet.</div>`;
                }
            });
        } catch (error) {
            console.error("Fatal error in populateTree:", error);
            if (error instanceof OfficeExtension.Error) {
                console.error("Debug info: " + JSON.stringify(error.debugInfo));
            }
            treeContainer.innerHTML = `<div class="explorer-error">Error: Could not render the settings tree. Check the console for details.</div>`;
        }
    };

    /**
     * Recursively builds the HTML for the settings tree.
     * @param {object} data - The settings object or sub-object to render.
     * @param {HTMLElement} parentElement - The UL element to append children to.
     */
    const buildTree = (data, parentElement) => {
        for (const key in data) {
            if (Object.prototype.hasOwnProperty.call(data, key)) {
                const value = data[key];
                const li = document.createElement('li');
                const nodeHeader = document.createElement('div');
                nodeHeader.className = 'node-header';

                const keySpan = document.createElement('span');
                keySpan.className = 'node-key';
                keySpan.textContent = key;
                
                if (value !== null && typeof value === 'object') {
                    // This is a collapsible node (object or array)
                    const isArray = Array.isArray(value);
                    li.className = 'collapsible-node expanded'; // Expand by default
                    nodeHeader.innerHTML = ICONS.chevron + (isArray ? ICONS.array : ICONS.object);
                    nodeHeader.appendChild(keySpan);
                    
                    if (isArray) {
                        const countSpan = document.createElement('span');
                        countSpan.className = 'node-array-count';
                        countSpan.textContent = `(${value.length} items)`;
                        keySpan.appendChild(countSpan);
                    }

                    const childUl = document.createElement('ul');
                    li.appendChild(nodeHeader);
                    li.appendChild(childUl);
                    
                    nodeHeader.addEventListener('click', (e) => {
                        e.stopPropagation();
                        li.classList.toggle('expanded');
                    });
                    
                    buildTree(value, childUl);
                } else {
                    // This is a simple value node (leaf)
                    const valueType = value === null ? 'null' : typeof value;
                    nodeHeader.innerHTML = ICONS[valueType] || ICONS.string; // Fallback to string icon
                    nodeHeader.appendChild(keySpan);
                    nodeHeader.appendChild(document.createTextNode(': '));
                    
                    const valueSpan = document.createElement('span');
                    valueSpan.className = `node-value-${valueType}`;
                    valueSpan.textContent = value;
                    
                    const copyButton = document.createElement('button');
                    copyButton.className = 'copy-button icon-button';
                    copyButton.title = 'Copy value';
                    copyButton.innerHTML = ICONS.copy;

                    copyButton.onclick = (e) => {
                        e.stopPropagation();
                        navigator.clipboard.writeText(value).then(() => {
                            copyButton.innerHTML = ICONS.check;
                            copyButton.classList.add('copied');
                            setTimeout(() => {
                                copyButton.innerHTML = ICONS.copy;
                                copyButton.classList.remove('copied');
                            }, 1500);
                        }).catch(err => {
                            console.error('Failed to copy text: ', err);
                        });
                    };

                    nodeHeader.appendChild(valueSpan);
                    nodeHeader.appendChild(copyButton);
                    li.appendChild(nodeHeader);
                }
                parentElement.appendChild(li);
            }
        }
    };

    // --- Attach Event Listeners ---
    console.log("Attaching event listener to close button.");
    // The messageParent API is the standard way to close a dialog in Office Add-ins.
    // We pass a message (can be anything) to signal the host page to close this dialog.
    closeButton.onclick = () => {
        console.log("Close button clicked. Sending 'close' message to parent.");
        Office.context.ui.messageParent("close");
    };

    // Initial population of the tree view
    populateTree();
}

