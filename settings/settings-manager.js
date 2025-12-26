/*
 * settings-manager.js
 * Version: 1.4.0
 *
 * Main settings management for the Student Retention Add-in.
 *
 * Handles:
 * - Settings UI initialization and population
 * - User management (add, edit, remove users)
 * - LDA report configuration
 * - Column management with drag-and-drop
 * - Settings persistence to Office document
 * - Accordion/collapsible sections
 */

let settings = {};

Office.onReady((info) => {
    if (info.host === Office.HostType.Excel) {
        // Initialize with current settings and populate the UI
        loadSettingsAndPopulateUI();

        // Add event listeners for the main buttons
        document.getElementById(CONSTANTS.DOM.IDS.saveButton).onclick = saveSettings;
        document.getElementById(CONSTANTS.DOM.IDS.resetButton).onclick = resetSettings;
        
        // Initialize the "Add User" modal functionality
        initializeAddUserModal(settings, saveSettings);
        
        // Initialize the collapsible sections
        initializeAccordions();
        
        // Initialize the settings explorer modal
        initializeSettingsExplorerModal();
    }
});

/**
 * Initializes the behavior for all collapsible sections on the page.
 * Sections are collapsed by default and can be toggled by clicking the header.
 */
function initializeAccordions() {
    document.querySelectorAll(".section-header.collapsible").forEach(header => {
        const content = header.nextElementSibling;

        // Set initial ARIA attributes for accessibility
        header.setAttribute('aria-expanded', 'false');
        if (content) {
            const contentId = `section-content-${Math.random().toString(36).substr(2, 9)}`;
            content.setAttribute('id', contentId);
            header.setAttribute('aria-controls', contentId);
        }

        const toggleSection = () => {
            const isExpanded = header.getAttribute('aria-expanded') === 'true';
            header.setAttribute('aria-expanded', !isExpanded);
            header.classList.toggle('expanded');
            if (content) {
                content.classList.toggle('expanded');
            }
        };

        header.addEventListener('click', (e) => {
            // Prevent the accordion from toggling if a button or other interactive element in the header is clicked.
            if (e.target.closest('button, input, select, a')) {
                return;
            }
            toggleSection();
        });

        header.addEventListener('keydown', (event) => {
            // Allow toggling with Enter or Space key for accessibility
            if (event.key === "Enter" || event.key === " ") {
                event.preventDefault();
                toggleSection();
            }
        });
    });
}

/**
 * Resets the add-in settings to their default values by removing the settings key
 * from the document and then reloading the UI.
 */
function resetSettings() {
    Office.context.document.settings.remove(CONSTANTS.SETTINGS_KEYS.APP);
    Office.context.document.settings.saveAsync(function (asyncResult) {
        const status = document.getElementById(CONSTANTS.DOM.IDS.status);
        if (asyncResult.status == Office.AsyncResultStatus.Failed) {
            console.error('Settings failed to reset. Error: ' + asyncResult.error.message);
            status.textContent = 'Error resetting settings.';
            status.className = `${CONSTANTS.DOM.CLASSES.statusMessage} ${CONSTANTS.DOM.CLASSES.statusError} ${CONSTANTS.DOM.CLASSES.visible}`;
        } else {
            console.log('Settings reset successfully.');
            status.textContent = 'Settings have been reset to default.';
            status.className = `${CONSTANTS.DOM.CLASSES.statusMessage} ${CONSTANTS.DOM.CLASSES.statusSuccess} ${CONSTANTS.DOM.CLASSES.visible}`;
            
            // Reload the UI to apply the default settings
            loadSettingsAndPopulateUI();
        }
        // Clear the status message after a few seconds
        setTimeout(() => {
            status.textContent = '';
            status.className = CONSTANTS.DOM.CLASSES.statusMessage;
        }, 3000);
    });
}

/**
 * Loads settings from the document, merges them with defaults to prevent errors,
 * and populates all UI elements on the settings page.
 */
function loadSettingsAndPopulateUI() {
    const settingsString = Office.context.document.settings.get(CONSTANTS.SETTINGS_KEYS.APP);
    if (settingsString) {
        try {
            // Merge saved settings with defaults to ensure all keys exist and handle future additions
            const savedSettings = JSON.parse(settingsString);
            settings = {
                ...JSON.parse(JSON.stringify(CONSTANTS.DEFAULT_SETTINGS)), // Deep copy of defaults
                ...savedSettings,
                createlda: { ...CONSTANTS.DEFAULT_SETTINGS.createlda, ...(savedSettings.createlda || {}) },
                userProfile: { ...CONSTANTS.DEFAULT_SETTINGS.userProfile, ...(savedSettings.userProfile || {}) },
                taskpane: { ...CONSTANTS.DEFAULT_SETTINGS.taskpane, ...(savedSettings.taskpane || {}) }
            };
        } catch (e) {
            console.error("Error parsing settings, reverting to defaults:", e);
            settings = JSON.parse(JSON.stringify(CONSTANTS.DEFAULT_SETTINGS)); // Deep copy on error
        }
    } else {
        // If no settings exist, initialize with a deep copy of the defaults
        settings = JSON.parse(JSON.stringify(CONSTANTS.DEFAULT_SETTINGS));
    }
    
    // Ensure userProfile object and userList array exist
    settings.userProfile = settings.userProfile || { ...CONSTANTS.DEFAULT_SETTINGS.userProfile };
    settings.userProfile.userList = settings.userProfile.userList || [];

    // Set the current user's name and add them to the list if not already present
    settings.userProfile.name = Office.context.displayName || "";
    if (settings.userProfile.name && !settings.userProfile.userList.includes(settings.userProfile.name)) {
        settings.userProfile.userList.push(settings.userProfile.name);
    }

    // --- Populate UI Elements ---
    document.getElementById(CONSTANTS.DOM.IDS.smartNavigationToggle).checked = settings.taskpane.smartNavigation;
    document.getElementById(CONSTANTS.DOM.IDS.daysOutFilter).value = settings.createlda.daysOutFilter;
    document.getElementById(CONSTANTS.DOM.IDS.includeFailingList).checked = settings.createlda.includeFailingList;
    document.getElementById(CONSTANTS.DOM.IDS.includeLdaTagFollowup).checked = settings.createlda.includeLdaTagFollowup;
    document.getElementById(CONSTANTS.DOM.IDS.hideLeftoverColumns).checked = settings.createlda.hideLeftoverColumns;
    document.getElementById(CONSTANTS.DOM.IDS.treatEmptyGradesAsZero).checked = settings.createlda.treatEmptyGradesAsZero;

    // Render the dynamic UI components
    renderUserList();
    loadAndRenderLdaColumns();
}

/**
 * Renders the list of users in the User Management section with interactive elements
 * for editing and removing users.
 */
function renderUserList() {
    const container = document.getElementById(CONSTANTS.DOM.IDS.userListContainer);
    container.innerHTML = ''; // Clear existing list to prevent duplicates
    const userList = settings.userProfile.userList || [];

    if (userList.length === 0) {
        container.innerHTML = '<p class="text-gray-500 text-sm">No users found. Add a user to get started.</p>';
        return;
    }

    userList.forEach(user => {
        const userItem = document.createElement('div');
        userItem.className = CONSTANTS.DOM.CLASSES.userItem;

        const nameContainer = document.createElement('div');
        nameContainer.className = CONSTANTS.DOM.CLASSES.userNameContainer;

        const nameDisplay = document.createElement('span');
        nameDisplay.textContent = user;
        nameDisplay.className = CONSTANTS.DOM.CLASSES.userNameDisplay;

        const nameInput = document.createElement('input');
        nameInput.type = 'text';
        nameInput.value = user;
        nameInput.className = `${CONSTANTS.DOM.CLASSES.userNameInput} ${CONSTANTS.DOM.CLASSES.hidden}`;

        nameContainer.append(nameDisplay, nameInput);

        const actions = document.createElement('div');
        actions.className = CONSTANTS.DOM.CLASSES.userItemActions;
        
        // --- Create Action Buttons ---
        const editButton = createIconButton('Edit User', 'edit-button', `<svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"></path><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"></path></svg>`);
        const saveButton = createIconButton('Save Changes', 'save-button', `<svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><polyline points="20 6 9 17 4 12"></polyline></svg>`, true);
        const cancelButton = createIconButton('Cancel Edit', 'cancel-button', `<svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><line x1="18" y1="6" x2="6" y2="18"></line><line x1="6" y1="6" x2="18" y2="18"></line></svg>`, true);
        const removeButton = createIconButton('Remove User', 'remove-button', `<svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><polyline points="3 6 5 6 21 6"></polyline><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"></path></svg>`);
        removeButton.disabled = userList.length <= 1;

        // --- Event Listeners for Edit Mode ---
        const toggleEditMode = (isEditing) => {
            nameDisplay.classList.toggle(CONSTANTS.DOM.CLASSES.hidden, isEditing);
            editButton.classList.toggle(CONSTANTS.DOM.CLASSES.hidden, isEditing);
            removeButton.classList.toggle(CONSTANTS.DOM.CLASSES.hidden, isEditing);
            nameInput.classList.toggle(CONSTANTS.DOM.CLASSES.hidden, !isEditing);
            saveButton.classList.toggle(CONSTANTS.DOM.CLASSES.hidden, !isEditing);
            cancelButton.classList.toggle(CONSTANTS.DOM.CLASSES.hidden, !isEditing);
            if (isEditing) {
                nameInput.focus();
                nameInput.select();
            }
        };

        editButton.onclick = () => toggleEditMode(true);
        cancelButton.onclick = () => {
            nameInput.value = user; // Reset input value
            toggleEditMode(false);
        };

        saveButton.onclick = () => {
            const oldName = user;
            const newName = nameInput.value.trim();

            if (newName && newName !== oldName) {
                if (settings.userProfile.userList.includes(newName)) {
                    console.error(`User "${newName}" already exists.`);
                    nameInput.style.borderColor = 'red';
                    setTimeout(() => { nameInput.style.borderColor = ''; }, 2000);
                    return;
                }
                const index = settings.userProfile.userList.indexOf(oldName);
                if (index > -1) {
                    settings.userProfile.userList[index] = newName;
                    if (settings.userProfile.name === oldName) {
                        settings.userProfile.name = newName;
                    }
                    saveSettings(); // This will re-render the list
                }
            } else {
                toggleEditMode(false); // If name is empty or unchanged, just cancel
            }
        };

        nameInput.onkeydown = (event) => {
            if (event.key === 'Enter') saveButton.onclick();
            else if (event.key === 'Escape') cancelButton.onclick();
        };

        removeButton.onclick = () => removeUser(user);

        actions.append(editButton, saveButton, cancelButton, removeButton);
        userItem.append(nameContainer, actions);
        container.appendChild(userItem);
    });
}

/**
 * Helper function to create an icon button for the user list actions.
 * @param {string} title - The tooltip text for the button.
 * @param {string} extraClass - An additional CSS class for the button.
 * @param {string} innerHTML - The SVG icon for the button.
 * @param {boolean} isHidden - Whether the button should be hidden initially.
 * @returns {HTMLButtonElement} The created button element.
 */
function createIconButton(title, extraClass, innerHTML, isHidden = false) {
    const button = document.createElement('button');
    button.className = `${CONSTANTS.DOM.CLASSES.iconButton} ${extraClass}`;
    if (isHidden) {
        button.classList.add(CONSTANTS.DOM.CLASSES.hidden);
    }
    button.title = title;
    button.innerHTML = innerHTML;
    return button;
}


/**
 * Removes a user from the settings and triggers a save and UI refresh.
 * @param {string} nameToRemove - The name of the user to be removed.
 */
function removeUser(nameToRemove) {
    const userList = settings.userProfile.userList;
    if (userList.length <= 1) {
        console.warn("Cannot remove the last user.");
        // Optionally, show a message to the user here.
        return;
    }
    settings.userProfile.userList = userList.filter(u => u !== nameToRemove);
    // If the removed user was the 'active' one, select the first user in the list.
    if (settings.userProfile.name === nameToRemove) {
        settings.userProfile.name = settings.userProfile.userList[0] || "";
    }
    saveSettings(); 
}


/**
 * Asynchronously loads the header columns from the "Master List" sheet and renders them
 * into the "Included" and "Available" columns lists for drag-and-drop management.
 */
async function loadAndRenderLdaColumns() {
    const includedContainer = document.getElementById(CONSTANTS.DOM.IDS.includedColumns);
    const availableContainer = document.getElementById(CONSTANTS.DOM.IDS.availableColumns);
    const loader = document.getElementById(CONSTANTS.DOM.IDS.columnsLoader);
    
    try {
        await Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getItem(CONSTANTS.SHEET_NAMES.MASTER_LIST);
            const headerRange = sheet.getRange("1:1").getUsedRange(true);
            headerRange.load("values");
            await context.sync();

            const allMasterColumns = headerRange.values[0].filter(h => h && String(h).trim() !== "");
            includedContainer.innerHTML = ""; // Clear previous items
            availableContainer.innerHTML = "";

            const selectedColumns = settings.createlda.ldaColumns || [];
            const selectedSet = new Set(selectedColumns);

            // Populate included columns based on the saved order
            selectedColumns.forEach(header => {
                if (allMasterColumns.includes(header)) {
                    includedContainer.appendChild(createColumnItem(header));
                }
            });

            // Populate available columns with all remaining (unselected) items
            allMasterColumns.forEach(header => {
                if (!selectedSet.has(header)) {
                    availableContainer.appendChild(createColumnItem(header));
                }
            });
            
            loader.style.display = 'none'; // Hide loader on success
        });

        // Initialize SortableJS to enable drag-and-drop between the two lists
        const sharedSortableOptions = {
            group: 'shared-columns',
            animation: 150,
            ghostClass: CONSTANTS.DOM.CLASSES.sortableGhost
        };
        new Sortable(includedContainer, sharedSortableOptions);
        new Sortable(availableContainer, sharedSortableOptions);

    } catch (error) {
        console.error("Error loading master list columns:", error);
        if (error instanceof OfficeExtension.Error) {
            console.error("Debug info:", JSON.stringify(error.debugInfo));
        }
        loader.innerHTML = `<p class="error-message">Error: Could not load columns. Make sure a sheet named '${CONSTANTS.SHEET_NAMES.MASTER_LIST}' exists and has a header row.</p>`;
        loader.classList.add(CONSTANTS.DOM.CLASSES.statusError, CONSTANTS.DOM.CLASSES.visible);
    }
}

/**
 * Creates a draggable column item element for the columns management section.
 * @param {string} header - The text content for the column item.
 * @returns {HTMLDivElement} The created column item element.
 */
function createColumnItem(header) {
    const colItem = document.createElement("div");
    colItem.className = CONSTANTS.DOM.CLASSES.columnItem;
    colItem.textContent = header;
    colItem.dataset.columnName = header; // Store the name in a data attribute
    return colItem;
}

/**
 * Gathers all current settings from the UI, updates the settings object,
 * and saves it to the document.
 */
function saveSettings() {
    // --- Gather Settings from UI ---
    settings.taskpane.smartNavigation = document.getElementById(CONSTANTS.DOM.IDS.smartNavigationToggle).checked;
    settings.createlda.daysOutFilter = parseInt(document.getElementById(CONSTANTS.DOM.IDS.daysOutFilter).value, 10) || 0;
    settings.createlda.includeFailingList = document.getElementById(CONSTANTS.DOM.IDS.includeFailingList).checked;
    settings.createlda.includeLdaTagFollowup = document.getElementById(CONSTANTS.DOM.IDS.includeLdaTagFollowup).checked;
    settings.createlda.hideLeftoverColumns = document.getElementById(CONSTANTS.DOM.IDS.hideLeftoverColumns).checked;
    settings.createlda.treatEmptyGradesAsZero = document.getElementById(CONSTANTS.DOM.IDS.treatEmptyGradesAsZero).checked;

    // Get the current order of columns from the "Included" list
    const includedContainer = document.getElementById(CONSTANTS.DOM.IDS.includedColumns);
    const selectedItems = includedContainer.querySelectorAll(`.${CONSTANTS.DOM.CLASSES.columnItem}`);
    settings.createlda.ldaColumns = Array.from(selectedItems).map(item => item.dataset.columnName);

    // --- Save to Document ---
    Office.context.document.settings.set(CONSTANTS.SETTINGS_KEYS.APP, JSON.stringify(settings));
    Office.context.document.settings.saveAsync(function (asyncResult) {
        const status = document.getElementById(CONSTANTS.DOM.IDS.status);
        if (asyncResult.status == Office.AsyncResultStatus.Failed) {
            console.error('Settings failed to save. Error: ' + asyncResult.error.message);
            status.textContent = 'Error saving settings.';
            status.className = `${CONSTANTS.DOM.CLASSES.statusMessage} ${CONSTANTS.DOM.CLASSES.statusError} ${CONSTANTS.DOM.CLASSES.visible}`;
        } else {
            console.log('Settings saved successfully.');
            status.textContent = 'Settings saved!';
            status.className = `${CONSTANTS.DOM.CLASSES.statusMessage} ${CONSTANTS.DOM.CLASSES.statusSuccess} ${CONSTANTS.DOM.CLASSES.visible}`;
            // Reload the entire UI to ensure consistency after saving
            loadSettingsAndPopulateUI();
        }
        // Clear the status message after a few seconds
        setTimeout(() => {
            status.textContent = '';
            status.className = CONSTANTS.DOM.CLASSES.statusMessage;
        }, 3000);
    });
}
