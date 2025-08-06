let settings = {};
const CONSTANTS = {
    MASTER_LIST_SHEET: "Master List",
    SETTINGS_KEY: "studentRetentionSettings"
};

Office.onReady((info) => {
    if (info.host === Office.HostType.Excel) {
        // Initialize with current settings and populate the UI
        loadSettingsAndPopulateUI();

        // Add event listeners
        document.getElementById("save-button").onclick = saveSettings;
        document.getElementById("reset-button").onclick = resetSettings;
        
        // --- Modal Event Listeners ---
        const addUserModal = document.getElementById("add-user-modal");
        const addUserButton = document.getElementById("add-user-button");
        const cancelAddUserButton = document.getElementById("cancel-add-user-button");
        const saveNewUserButton = document.getElementById("save-new-user-button");
        const firstNameInput = document.getElementById("new-user-first-name");
        const lastNameInput = document.getElementById("new-user-last-name");
        const addUserStatus = document.getElementById("add-user-status");

        const showModal = () => {
            addUserStatus.textContent = '';
            addUserStatus.className = 'status-message';
            firstNameInput.value = '';
            lastNameInput.value = '';
            addUserModal.style.display = 'flex';
            firstNameInput.focus();
        };

        const hideModal = () => {
            addUserModal.style.display = 'none';
        };

        addUserButton.onclick = showModal;
        cancelAddUserButton.onclick = hideModal;
        // Also hide modal if user clicks the overlay background
        addUserModal.addEventListener('click', (event) => {
            if (event.target === addUserModal) {
                hideModal();
            }
        });

        saveNewUserButton.onclick = () => {
            const firstName = firstNameInput.value.trim();
            const lastName = lastNameInput.value.trim();

            if (!firstName || !lastName) {
                addUserStatus.textContent = "Both names are required.";
                addUserStatus.className = 'status-message status-error visible';
                return;
            }

            const formattedName = `${firstName} ${lastName}`;

            if (settings.userProfile.userList.includes(formattedName)) {
                addUserStatus.textContent = `User "${formattedName}" already exists.`;
                addUserStatus.className = 'status-message status-error visible';
                return;
            }

            settings.userProfile.userList.push(formattedName);
            saveSettings(); // This will save and reload the UI
            hideModal();
        };
    }
});

function resetSettings() {
    // This function will remove the settings key from the document,
    // then reload the UI which will populate it with the default values.
    Office.context.document.settings.remove(CONSTANTS.SETTINGS_KEY);
    Office.context.document.settings.saveAsync(function (asyncResult) {
        const status = document.getElementById('status');
        if (asyncResult.status == Office.AsyncResultStatus.Failed) {
            console.log('Settings failed to reset. Error: ' + asyncResult.error.message);
            status.textContent = 'Error resetting settings.';
            status.className = 'status-message status-error visible';
        } else {
            console.log('Settings reset successfully.');
            status.textContent = 'Settings have been reset to default.';
            status.className = 'status-message status-success visible';
            
            // Reload the UI with default settings
            loadSettingsAndPopulateUI();
        }
        // Clear the message after a few seconds
        setTimeout(() => {
            status.textContent = '';
            status.className = 'status-message';
        }, 3000);
    });
}


function loadSettingsAndPopulateUI() {
    const settingsString = Office.context.document.settings.get(CONSTANTS.SETTINGS_KEY);
    if (settingsString) {
        try {
            settings = JSON.parse(settingsString);
        } catch (e) {
            console.error("Error parsing settings:", e);
            settings = {}; // Reset to avoid issues
        }
    } else {
        // If no settings string, initialize with empty object
        settings = {};
    }
    
    // Ensure settings objects exist with defaults
    if (!settings.createlda) {
        settings.createlda = {
            daysOutFilter: 6,
            includeFailingList: true,
            hideLeftoverColumns: true,
            treatEmptyGradesAsZero: false,
            ldaColumns: ['Assigned', 'StudentName', 'StudentNumber', 'LDA', 'Days Out', 'grade', 'Phone', 'Outreach']
        };
    }
    if (!settings.userProfile) {
        settings.userProfile = {
            name: Office.context.displayName || "",
            userList: []
        };
    }
    if (!settings.userProfile.userList) {
        settings.userProfile.userList = [];
    }
    if (!settings.taskpane) {
        settings.taskpane = {
            smartNavigation: true // Default to true
        };
    }
    
    // Add current user to the list if not already present
    const currentName = settings.userProfile.name || Office.context.displayName;
    if (currentName && !settings.userProfile.userList.includes(currentName)) {
        settings.userProfile.userList.push(currentName);
    }

    // Populate UI for Task Pane settings
    document.getElementById("smart-navigation-toggle").checked = settings.taskpane.smartNavigation !== false;

    // Populate UI for LDA Report settings
    document.getElementById("days-out-filter").value = settings.createlda.daysOutFilter || 6;
    document.getElementById("include-failing-list").checked = settings.createlda.includeFailingList !== false;
    document.getElementById("hide-leftover-columns").checked = settings.createlda.hideLeftoverColumns !== false;
    document.getElementById("treat-empty-grades-as-zero").checked = settings.createlda.treatEmptyGradesAsZero === true;

    // Render the user list management UI
    renderUserList();
    // Load and render the LDA column selector
    loadAndRenderLdaColumns();
}

function renderUserList() {
    const container = document.getElementById('user-list-container');
    container.innerHTML = ''; // Clear existing list
    const userList = (settings.userProfile && settings.userProfile.userList) || [];
    const activeUser = settings.userProfile.name || Office.context.displayName;

    if (userList.length === 0) {
        container.innerHTML = '<p class="text-gray-500 text-sm">No users found. Save settings to add the current user.</p>';
        return;
    }

    userList.forEach(user => {
        const userItem = document.createElement('div');
        userItem.className = 'user-item';
        
        const radio = document.createElement('input');
        radio.type = 'radio';
        radio.name = 'active-user';
        radio.value = user;
        radio.id = `user-radio-${user.replace(/\s+/g, '-')}`;
        radio.checked = (user === activeUser);
        
        const label = document.createElement('label');
        label.htmlFor = radio.id;
        label.textContent = user;

        const actions = document.createElement('div');
        actions.className = 'user-item-actions';
        
        const editButton = document.createElement('button');
        editButton.className = 'icon-button';
        editButton.innerHTML = `<svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"></path><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"></path></svg>`;
        editButton.onclick = () => editUser(user);
        
        const removeButton = document.createElement('button');
        removeButton.className = 'icon-button';
        removeButton.innerHTML = `<svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><polyline points="3 6 5 6 21 6"></polyline><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"></path></svg>`;
        removeButton.onclick = () => removeUser(user);
        removeButton.disabled = (user === activeUser);

        userItem.appendChild(radio);
        userItem.appendChild(label);
        actions.appendChild(editButton);
        actions.appendChild(removeButton);
        userItem.appendChild(actions);
        container.appendChild(userItem);
    });
}

function editUser(oldName) {
    // For simplicity, we'll still use prompt here, but this could also be converted to a modal.
    const newName = prompt("Enter the new name for this user:", oldName);
    if (newName && newName.trim() !== "" && newName !== oldName) {
        const userList = settings.userProfile.userList;
        const index = userList.indexOf(oldName);
        if (index > -1) {
            userList[index] = newName;
            // If the edited user was the active user, update the active name
            if (settings.userProfile.name === oldName) {
                settings.userProfile.name = newName;
            }
            saveSettings(); // Save and reload the UI
        }
    }
}

function removeUser(nameToRemove) {
    // `confirm` is also not supported, so we'll just remove without confirmation for now.
    // A proper solution would be a custom confirmation modal.
    settings.userProfile.userList = settings.userProfile.userList.filter(u => u !== nameToRemove);
    saveSettings(); // Save and reload the UI
}


async function loadAndRenderLdaColumns() {
    const includedContainer = document.getElementById("included-columns");
    const availableContainer = document.getElementById("available-columns");
    const loader = document.getElementById("columns-loader");
    
    try {
        await Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getItem(CONSTANTS.MASTER_LIST_SHEET);
            const headerRange = sheet.getRange("1:1").getUsedRange(true);
            headerRange.load("values");
            await context.sync();

            const allMasterColumns = headerRange.values[0].filter(h => h && String(h).trim() !== "");
            includedContainer.innerHTML = ""; // Clear previous items
            availableContainer.innerHTML = "";

            const selectedColumns = settings.createlda.ldaColumns || [];
            const selectedSet = new Set(selectedColumns);

            // Populate included columns in the saved order
            selectedColumns.forEach(header => {
                if (allMasterColumns.includes(header)) {
                    includedContainer.appendChild(createColumnItem(header));
                }
            });

            // Populate available columns with the remaining items
            allMasterColumns.forEach(header => {
                if (!selectedSet.has(header)) {
                    availableContainer.appendChild(createColumnItem(header));
                }
            });
            
            loader.style.display = 'none';
        });

        // Initialize SortableJS on both lists
        const sharedSortableOptions = {
            group: 'shared-columns', // set both lists to same group
            animation: 150,
            ghostClass: 'sortable-ghost'
        };
        new Sortable(includedContainer, sharedSortableOptions);
        new Sortable(availableContainer, sharedSortableOptions);

    } catch (error) {
        console.error("Error loading master list columns:", error);
        if (error instanceof OfficeExtension.Error) {
            console.error("Debug info:", JSON.stringify(error.debugInfo));
        }
        loader.innerHTML = `<p class="error-message">Error: Could not load columns. Make sure a sheet named 'Master List' exists and has a header row.</p>`;
        loader.classList.add('status-error', 'visible');
    }
}

function createColumnItem(header) {
    const colItem = document.createElement("div");
    colItem.className = "column-item";
    colItem.textContent = header;
    colItem.dataset.columnName = header;
    return colItem;
}


function saveSettings() {
    // --- User Profile ---
    // The user list is modified directly by the edit/remove functions.
    // We just need to find the selected active user.
    const selectedRadio = document.querySelector('input[name="active-user"]:checked');
    if (selectedRadio) {
        settings.userProfile.name = selectedRadio.value;
    }

    // --- Task Pane ---
    settings.taskpane.smartNavigation = document.getElementById("smart-navigation-toggle").checked;

    // --- LDA Report ---
    settings.createlda.daysOutFilter = parseInt(document.getElementById("days-out-filter").value, 10);
    settings.createlda.includeFailingList = document.getElementById("include-failing-list").checked;
    settings.createlda.hideLeftoverColumns = document.getElementById("hide-leftover-columns").checked;
    settings.createlda.treatEmptyGradesAsZero = document.getElementById("treat-empty-grades-as-zero").checked;

    // Get selected columns from the "Included" list, preserving their order
    const includedContainer = document.getElementById("included-columns");
    const selectedItems = includedContainer.querySelectorAll(".column-item");
    settings.createlda.ldaColumns = Array.from(selectedItems).map(item => item.dataset.columnName);

    // Save the updated settings object
    Office.context.document.settings.set(CONSTANTS.SETTINGS_KEY, JSON.stringify(settings));
    Office.context.document.settings.saveAsync(function (asyncResult) {
        const status = document.getElementById('status');
        if (asyncResult.status == Office.AsyncResultStatus.Failed) {
            console.log('Settings failed to save. Error: ' + asyncResult.error.message);
            status.textContent = 'Error saving settings.';
            status.className = 'status-message status-error visible';
        } else {
            console.log('Settings saved successfully.');
            status.textContent = 'Settings saved!';
            status.className = 'status-message status-success visible';
            // Reload UI to reflect all changes consistently
            loadSettingsAndPopulateUI();
        }
        // Clear the message after a few seconds
        setTimeout(() => {
            status.textContent = '';
            status.className = 'status-message';
        }, 3000);
    });
}
