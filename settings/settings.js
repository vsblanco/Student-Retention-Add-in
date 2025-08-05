let settings = {};
let userBeforeEdit = ""; // To track the original name being edited
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
        document.getElementById("current-user-select").onchange = handleUserSwitch;
    }
});

function handleUserSwitch() {
    const dropdown = document.getElementById("current-user-select");
    const editInput = document.getElementById("edit-user-name");
    const selectedUser = dropdown.value;
    
    editInput.value = selectedUser;
    userBeforeEdit = selectedUser; // Update the user being edited
    renderUserList(); // Re-render to update the "Current User" badge
}


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
            treatEmptyGradesAsZero: false, // Default to false
            ldaColumns: ['Assigned', 'StudentName', 'StudentNumber', 'LDA', 'Days Out', 'grade', 'Phone', 'Outreach']
        };
    }
    if (!settings.userProfile) {
        settings.userProfile = {
            name: Office.context.displayName || "",
            userList: [] // Initialize user list
        };
    }
    // Handle older settings objects that might not have userList
    if (!settings.userProfile.userList) {
        settings.userProfile.userList = [];
    }
    
    // Add current user to the list if not already present
    const currentName = settings.userProfile.name || Office.context.displayName;
    if (currentName && !settings.userProfile.userList.includes(currentName)) {
        settings.userProfile.userList.push(currentName);
    }

    userBeforeEdit = currentName; // Store the name for editing purposes

    // Populate UI for User Profile
    const userDropdown = document.getElementById("current-user-select");
    userDropdown.innerHTML = '';
    settings.userProfile.userList.forEach(user => {
        const option = document.createElement('option');
        option.value = user;
        option.textContent = user;
        if (user === currentName) {
            option.selected = true;
        }
        userDropdown.appendChild(option);
    });
    document.getElementById("edit-user-name").value = currentName;


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
    const currentUser = document.getElementById("current-user-select").value;

    if (userList.length === 0) {
        container.innerHTML = '<p class="text-gray-500 text-sm">No users found.</p>';
        return;
    }

    userList.forEach(user => {
        const userItem = document.createElement('div');
        userItem.className = 'setting-item';
        
        const userName = document.createElement('span');
        userName.textContent = user;
        userItem.appendChild(userName);

        if (user !== currentUser) {
            const deleteButton = document.createElement('button');
            deleteButton.textContent = 'Remove';
            deleteButton.className = 'button-danger-small'; // A new smaller button style
            deleteButton.onclick = () => {
                // Remove user from settings and re-render
                settings.userProfile.userList = settings.userProfile.userList.filter(u => u !== user);
                loadSettingsAndPopulateUI(); // Reload everything to ensure consistency
            };
            userItem.appendChild(deleteButton);
        } else {
             const currentUserBadge = document.createElement('span');
             currentUserBadge.textContent = 'Current User';
             currentUserBadge.className = 'current-user-badge'; // A new badge style
             userItem.appendChild(currentUserBadge);
        }

        container.appendChild(userItem);
    });
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
    const editedName = document.getElementById("edit-user-name").value.trim();
    const selectedUser = document.getElementById("current-user-select").value;

    // If the name was edited, update it in the user list
    if (editedName && editedName !== userBeforeEdit) {
        const userList = settings.userProfile.userList || [];
        const index = userList.indexOf(userBeforeEdit);
        if (index > -1) {
            userList[index] = editedName;
        }
        // If the edited user was the current user, update the current user name
        if (settings.userProfile.name === userBeforeEdit) {
            settings.userProfile.name = editedName;
        }
    }
    
    // Set the active user based on the dropdown selection
    settings.userProfile.name = selectedUser;
    // If the name was edited, the dropdown value might be stale, so we find the new name
    if (editedName && editedName !== userBeforeEdit && selectedUser === userBeforeEdit) {
        settings.userProfile.name = editedName;
    }


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
