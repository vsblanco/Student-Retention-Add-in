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
            treatEmptyGradesAsZero: false, // Default to false
            ldaColumns: ['Assigned', 'StudentName', 'StudentNumber', 'LDA', 'Days Out', 'grade', 'Phone', 'Outreach']
        };
    }
    if (!settings.userProfile) {
        settings.userProfile = {
            name: Office.context.displayName || ""
        };
    }

    // Populate UI for User Profile
    document.getElementById("user-full-name").value = settings.userProfile.name || "";

    // Populate UI for LDA Report settings
    document.getElementById("days-out-filter").value = settings.createlda.daysOutFilter || 6;
    document.getElementById("include-failing-list").checked = settings.createlda.includeFailingList !== false;
    document.getElementById("hide-leftover-columns").checked = settings.createlda.hideLeftoverColumns !== false;
    document.getElementById("treat-empty-grades-as-zero").checked = settings.createlda.treatEmptyGradesAsZero === true;


    // Load and render the LDA column selector
    loadAndRenderLdaColumns();
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
    // Get values from the UI
    // User Profile
    settings.userProfile.name = document.getElementById("user-full-name").value.trim();

    // LDA Report
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
        }
        // Clear the message after a few seconds
        setTimeout(() => {
            status.textContent = '';
            status.className = 'status-message';
        }, 3000);
    });
}
