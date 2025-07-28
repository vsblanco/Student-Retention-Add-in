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
    }
});

function loadSettingsAndPopulateUI() {
    const settingsString = Office.context.document.settings.get(CONSTANTS.SETTINGS_KEY);
    if (settingsString) {
        try {
            settings = JSON.parse(settingsString);
        } catch (e) {
            console.error("Error parsing settings:", e);
            settings = {}; // Reset to avoid issues
        }
    }
    
    // Ensure settings object and createlda property exist with defaults
    if (!settings.createlda) {
        settings.createlda = {
            daysOutFilter: 6,
            includeFailingList: true,
            ldaColumns: ['Assigned', 'StudentName', 'StudentNumber', 'LDA', 'Days Out', 'grade', 'Phone', 'Outreach']
        };
    }

    // Populate UI for existing fields
    document.getElementById("days-out-filter").value = settings.createlda.daysOutFilter || 6;
    document.getElementById("include-failing-list").checked = settings.createlda.includeFailingList !== false;

    // Load and render the new column selector
    loadAndRenderLdaColumns();
}

async function loadAndRenderLdaColumns() {
    const container = document.getElementById("lda-columns-container");
    try {
        await Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getItem(CONSTANTS.MASTER_LIST_SHEET);
            const headerRange = sheet.getRange("1:1").getUsedRange(true);
            headerRange.load("values");
            await context.sync();

            const headers = headerRange.values[0].filter(h => h && String(h).trim() !== "");
            container.innerHTML = ""; // Clear previous items

            const selectedColumns = new Set(settings.createlda.ldaColumns || []);

            headers.forEach(header => {
                const colItem = document.createElement("div");
                colItem.className = "column-item";
                colItem.textContent = header;
                colItem.dataset.columnName = header;

                if (selectedColumns.has(header)) {
                    colItem.classList.add("selected");
                }

                colItem.onclick = () => {
                    colItem.classList.toggle("selected");
                };

                container.appendChild(colItem);
            });
        });
    } catch (error) {
        console.error("Error loading master list columns:", error);
        if (error instanceof OfficeExtension.Error) {
            console.error("Debug info:", JSON.stringify(error.debugInfo));
        }
        container.innerHTML = `<p class="error-message">Error: Could not load columns. Make sure a sheet named 'Master List' exists and has a header row.</p>`;
    }
}

function saveSettings() {
    // Get values from the UI
    settings.createlda.daysOutFilter = parseInt(document.getElementById("days-out-filter").value, 10);
    settings.createlda.includeFailingList = document.getElementById("include-failing-list").checked;

    // Get selected columns from the new UI
    const container = document.getElementById("lda-columns-container");
    const selectedItems = container.querySelectorAll(".column-item.selected");
    settings.createlda.ldaColumns = Array.from(selectedItems).map(item => item.dataset.columnName);

    // Save the updated settings object
    Office.context.document.settings.set(CONSTANTS.SETTINGS_KEY, JSON.stringify(settings));
    Office.context.document.settings.saveAsync(function (asyncResult) {
        const status = document.getElementById('status');
        if (asyncResult.status == Office.AsyncResultStatus.Failed) {
            console.log('Settings failed to save. Error: ' + asyncResult.error.message);
            status.textContent = 'Error saving settings.';
            status.className = 'status-message status-error';
        } else {
            console.log('Settings saved successfully.');
            status.textContent = 'Settings saved!';
            status.className = 'status-message status-success';
        }
        // Clear the message after a few seconds
        setTimeout(() => {
            status.textContent = '';
            status.className = '';
        }, 3000);
    });
}
