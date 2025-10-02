// Timestamp: 2025-10-02 02:35 PM | Version: 1.2.0
/**
 * This script handles the functionality for the "Workbook Settings Explorer" modal,
 * which provides a read-only, collapsible tree view of all the add-in's settings
 * stored in the current workbook.
 */

// A simple icon registry for consistency.
const ICONS = {
    chevronRight: `<svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" class="chevron-icon"><polyline points="9 18 15 12 9 6"></polyline></svg>`,
    folder: `<svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M22 19a2 2 0 0 1-2 2H4a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h5l2 3h9a2 2 0 0 1 2 2z"></path></svg>`,
    file: `<svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M13 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V9z"></path><polyline points="13 2 13 9 20 9"></polyline></svg>`
};

/**
 * Initializes the settings explorer modal, attaching event listeners to the
 * open and close buttons.
 */
function initializeSettingsExplorerModal() {
    const viewButton = document.getElementById('view-settings-button');
    const modal = document.getElementById('settings-explorer-modal');
    const closeButton = document.getElementById('close-settings-explorer-button');
    const treeContainer = document.getElementById('settings-tree-container');

    viewButton.addEventListener('click', () => {
        // Fetch the latest settings from the document
        const settingsString = Office.context.document.settings.get('studentRetentionAppSettings');
        let settingsObject = {};
        try {
            settingsObject = settingsString ? JSON.parse(settingsString) : { "status": "No settings found in this workbook." };
        } catch (e) {
            console.error("Error parsing settings:", e);
            settingsObject = { "error": "Could not parse settings. They might be corrupted." };
        }

        // Build and render the tree view
        treeContainer.innerHTML = buildTree(settingsObject);
        
        // Show the modal
        modal.style.display = 'flex';
    });

    const closeModal = () => {
        modal.style.display = 'none';
        treeContainer.innerHTML = ''; // Clear content to save memory
    };

    closeButton.addEventListener('click', closeModal);

    // Close the modal if the user clicks on the overlay
    modal.addEventListener('click', (event) => {
        if (event.target === modal) {
            closeModal();
        }
    });

    // Add event delegation for toggling tree nodes
    treeContainer.addEventListener('click', (event) => {
        const nodeContent = event.target.closest('.tree-node-content');
        if (nodeContent) {
            const parentNode = nodeContent.parentElement;
            if (parentNode.classList.contains('expandable')) {
                parentNode.classList.toggle('expanded');
            }
        }
    });

    treeContainer.addEventListener('keydown', (event) => {
        const nodeContent = event.target.closest('.tree-node-content');
        if (nodeContent && (event.key === 'Enter' || event.key === ' ')) {
            event.preventDefault();
            const parentNode = nodeContent.parentElement;
            if (parentNode.classList.contains('expandable')) {
                parentNode.classList.toggle('expanded');
            }
        }
    });
}

/**
 * Recursively builds the HTML for the settings tree.
 * @param {object} obj - The settings object or sub-object to render.
 * @param {number} level - The current depth level for indentation.
 * @returns {string} The HTML string for the tree.
 */
function buildTree(obj, level = 0) {
    let html = '<ul class="settings-tree">';

    for (const key in obj) {
        if (obj.hasOwnProperty(key)) {
            const value = obj[key];
            const isObject = value !== null && typeof value === 'object';
            const hasChildren = isObject && Object.keys(value).length > 0;

            html += `<li class="tree-node ${hasChildren ? 'expandable' : ''}" style="--level: ${level};">`;
            html += `<div class="tree-node-content" tabindex="0">`;
            
            html += `<span class="chevron-container">${hasChildren ? ICONS.chevronRight : ''}</span>`;
            
            const icon = hasChildren ? ICONS.folder : ICONS.file;
            html += `<span class="icon-container">${icon}</span>`;

            html += `<span class="key">${key}:</span>`;

            if (!isObject) {
                html += `<span class="value">${formatValue(value)}</span>`;
            }

            html += `</div>`;

            if (hasChildren) {
                html += buildTree(value, level + 1);
            }

            html += `</li>`;
        }
    }

    html += '</ul>';
    return html;
}

/**
 * Formats a value for display, adding quotes for strings.
 * @param {*} value - The value to format.
 * @returns {string} The formatted value as an HTML string.
 */
function formatValue(value) {
    if (typeof value === 'string') {
        // Basic HTML escaping for display
        const escapedValue = value.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
        return `"${escapedValue}"`;
    }
    if (value === null) {
        return 'null';
    }
    return value.toString();
}

