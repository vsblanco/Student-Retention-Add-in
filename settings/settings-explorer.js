// Timestamp: 2025-10-02 10:37 AM | Version: 1.1.0

/**
 * Initializes the settings explorer modal, handling its opening, closing,
 * and the rendering of the settings data into a tree view.
 */
function initializeSettingsExplorerModal() {
    // --- Get DOM Elements ---
    const explorerModal = document.getElementById('settings-explorer-modal');
    const openButton = document.getElementById('view-settings-button');
    const closeButton = document.getElementById('close-settings-explorer-button');
    const treeContainer = document.getElementById('settings-tree-container');

    // --- SVG Icons ---
    const ICONS = {
        chevron: `<svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" class="tree-icon chevron"><polyline points="9 18 15 12 9 6"></polyline></svg>`,
        folder: `<svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" class="tree-icon"><path d="M22 19a2 2 0 0 1-2 2H4a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h5l2 3h9a2 2 0 0 1 2 2z"></path></svg>`,
        file: `<svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" class="tree-icon"><path d="M13 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V9z"></path><polyline points="13 2 13 9 20 9"></polyline></svg>`
    };

    // --- Modal Visibility ---
    const showModal = () => {
        populateTree();
        explorerModal.style.display = 'flex';
    };

    const hideModal = () => {
        explorerModal.style.display = 'none';
    };

    // --- Tree Building Logic ---

    /**
     * Fetches the latest settings and populates the tree view.
     */
    const populateTree = () => {
        treeContainer.innerHTML = ''; // Clear previous content
        const settingsString = Office.context.document.settings.get(CONSTANTS.SETTINGS_KEYS.APP);

        if (settingsString) {
            try {
                const settingsObject = JSON.parse(settingsString);
                const rootUl = document.createElement('ul');
                rootUl.className = 'settings-tree';
                buildTree(settingsObject, rootUl);
                treeContainer.appendChild(rootUl);
            } catch (error) {
                console.error("Error parsing settings for explorer:", error);
                treeContainer.textContent = 'Error: Could not parse settings JSON.';
            }
        } else {
            treeContainer.textContent = 'No settings have been saved for this add-in yet.';
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
                    li.className = 'collapsible-node';
                    const childUl = document.createElement('ul');
                    
                    // Add icons for folder/array
                    nodeHeader.innerHTML = ICONS.chevron;
                    nodeHeader.innerHTML += ICONS.folder;

                    // Add click listener to the header to toggle collapse
                    nodeHeader.addEventListener('click', (e) => {
                        e.stopPropagation(); // prevent event bubbling
                        li.classList.toggle('expanded');
                    });
                    
                    nodeHeader.appendChild(keySpan);
                    li.appendChild(nodeHeader);
                    li.appendChild(childUl);

                    // Recursively build the tree for the nested object/array
                    buildTree(value, childUl);
                } else {
                    // This is a simple value node (leaf)
                    nodeHeader.innerHTML = ICONS.file;
                    nodeHeader.appendChild(keySpan);
                    nodeHeader.appendChild(document.createTextNode(': '));
                    
                    const valueSpan = document.createElement('span');
                    const valueType = value === null ? 'null' : typeof value;
                    valueSpan.className = `node-value-${valueType}`;
                    valueSpan.textContent = value;
                    nodeHeader.appendChild(valueSpan);
                    li.appendChild(nodeHeader);
                }
                parentElement.appendChild(li);
            }
        }
    };

    // --- Attach Event Listeners ---
    openButton.onclick = showModal;
    closeButton.onclick = hideModal;

    // Hide the modal if the user clicks the dark overlay background
    explorerModal.addEventListener('click', (event) => {
        if (event.target === explorerModal) {
            hideModal();
        }
    });

    // Hide the modal if the user presses the Escape key
    document.addEventListener('keydown', (event) => {
        if (event.key === 'Escape' && explorerModal.style.display === 'flex') {
            hideModal();
        }
    });
}

