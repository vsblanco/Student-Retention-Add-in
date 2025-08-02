let settings = {};
const CONSTANTS = {
    MASTER_LIST_SHEET: "Master List",
    DOC_SETTINGS_KEY: "studentRetentionDocSettings",
    USER_SETTINGS_KEY: "studentRetentionUserSettings"
};

Office.onReady((info) => {
    if (info.host === Office.HostType.Excel) {
        // Initialize with current settings and populate the UI
        loadSettingsAndPopulateUI();

        // Add event listeners
        document.getElementById("save-button").onclick = saveSettings;
    }
});

function loadSettingsAndPopulateUI() {
    // Load document-specific settings (shared)
    const docSettingsString = Office.context.document.settings.get(CONSTANTS.DOC_SETTINGS_KEY);
    if (docSettingsString) {
        try {
            settings.createlda = JSON.parse(docSettingsString).createlda;
        } catch (e) {
            console.error("Error parsing document settings:", e);
        }
    }
    
    const finalizeSetup = () => {
        // This part runs after attempting to load roaming settings or if they are unavailable
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
    };

    // Load user-specific settings (per-user) from roaming settings, if available
    if (Office.context.roamingSettings) {
        Office.context.roamingSettings.refreshAsync(function(asyncResult) {
            if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
                const userSettingsString = Office.context.roamingSettings.get(CONSTANTS.USER_SETTINGS_KEY);
                if (userSettingsString) {
                    try {
                        settings.userProfile = JSON.parse(userSettingsString);
                    } catch (e) {
                        console.error("Error parsing user settings:", e);
                    }
                }
            } else {
                console.error("Failed to refresh roaming settings: " + asyncResult.error.message);
            }
            finalizeSetup(); // Call finalize here after async operation
        });
    } else {
        console.warn("Roaming settings not supported in this environment.");
        finalizeSetup(); // Call finalize here if roaming settings are not available
    }
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
    // --- Gather Document-Specific Settings ---
    const createldaSettings = {
        daysOutFilter: parseInt(document.getElementById("days-out-filter").value, 10),
        includeFailingList: document.getElementById("include-failing-list").checked,
        hideLeftoverColumns: document.getElementById("hide-leftover-columns").checked,
        treatEmptyGradesAsZero: document.getElementById("treat-empty-grades-as-zero").checked,
        ldaColumns: Array.from(document.querySelectorAll("#included-columns .column-item")).map(item => item.dataset.columnName)
    };
    const docSettings = { createlda: createldaSettings };
    
    // --- Gather User-Specific Settings ---
    const userProfileSettings = {
        ...settings.userProfile, // Preserve existing properties like hasSeenWelcomeMessage
        name: document.getElementById("user-full-name").value.trim()
    };

    // --- Save to respective locations ---
    const savePromises = [];

    // Always save document settings
    Office.context.document.settings.set(CONSTANTS.DOC_SETTINGS_KEY, JSON.stringify(docSettings));
    savePromises.push(new Promise(resolve => Office.context.document.settings.saveAsync(resolve)));

    // Conditionally save user settings if roaming settings are available
    if (Office.context.roamingSettings) {
        Office.context.roamingSettings.set(CONSTANTS.USER_SETTINGS_KEY, JSON.stringify(userProfileSettings));
        savePromises.push(new Promise(resolve => Office.context.roamingSettings.saveAsync(resolve)));
    } else {
        console.warn("Roaming settings not supported. User profile settings will not be saved.");
    }

    // Save both and provide a single status update
    Promise.all(savePromises).then(results => {
        const status = document.getElementById('status');
        const allSucceeded = results.every(res => res.status === Office.AsyncResultStatus.Succeeded);

        if (allSucceeded) {
            console.log('All settings saved successfully.');
            status.textContent = 'Settings saved!';
            status.className = 'status-message status-success visible';
        } else {
            console.log('One or more settings failed to save.');
            results.forEach(res => {
                if (res.status !== Office.AsyncResultStatus.Succeeded) {
                    console.error('Settings error: ' + res.error.message);
                }
            });
            status.textContent = 'Error saving settings.';
            status.className = 'status-message status-error visible';
        }
        
        setTimeout(() => {
            status.textContent = '';
            status.className = 'status-message';
        }, 3000);
    });
}
