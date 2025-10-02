// Timestamp: 2025-10-02 10:25 AM | Version: 1.0.0

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
                
                nodeHeader.appendChild(keySpan);
                
                if (value !== null && typeof value === 'object') {
                    // This is a collapsible node (object or array)
                    li.className = 'collapsible-node';
                    const childUl = document.createElement('ul');
                    
                    // Add click listener to the header to toggle collapse
                    nodeHeader.addEventListener('click', () => {
                        li.classList.toggle('expanded');
                    });
                    
                    li.appendChild(nodeHeader);
                    li.appendChild(childUl);

                    // Recursively build the tree for the nested object/array
                    buildTree(value, childUl);
                } else {
                    // This is a simple value node
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
