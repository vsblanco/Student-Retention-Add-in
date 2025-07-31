Office.onReady((info) => {
    if (info.host === Office.HostType.Excel) {
        document.addEventListener("DOMContentLoaded", run);
    }
});

const CONSTANTS = {
    SETTINGS_KEY: "studentRetentionSettings",
    COLUMN_MAPPINGS: {
        daysOut: ["days out", "daysout"],
    }
};

async function run() {
    setupTabs();
    try {
        await loadAnalytics();
    } catch (error) {
        showError(error.message);
        console.error(error);
    }
}

function setupTabs() {
    const tabCurrent = document.getElementById('tab-current');
    const tabProjection = document.getElementById('tab-projection');
    const panelCurrent = document.getElementById('panel-current');
    const panelProjection = document.getElementById('panel-projection');

    const tabs = [
        { button: tabCurrent, panel: panelCurrent },
        { button: tabProjection, panel: panelProjection }
    ];

    tabs.forEach(tab => {
        if (tab.button) {
            tab.button.addEventListener('click', () => {
                // Deactivate all tabs
                tabs.forEach(item => {
                    if (item.button) {
                        item.button.classList.remove('active', 'border-blue-500', 'text-blue-600');
                        item.button.classList.add('border-transparent', 'text-gray-500', 'hover:text-gray-700', 'hover:border-gray-300');
                    }
                    if (item.panel) {
                       item.panel.classList.add('hidden');
                    }
                });

                // Activate the clicked tab
                if (tab.button) {
                    tab.button.classList.add('active', 'border-blue-500', 'text-blue-600');
                    tab.button.classList.remove('border-transparent', 'text-gray-500');
                }
                if (tab.panel) {
                    tab.panel.classList.remove('hidden');
                }
            });
        }
    });
}

async function loadAnalytics() {
    const loadingMessage = document.getElementById("loading-message");
    const analyticsContent = document.getElementById("analytics-content");
    
    try {
        const [totalStudents, ldaStudents, projectedStudents] = await Promise.all([
            getTotalStudentCount(),
            getLdaStudentCount(),
            getProjectedLdaStudentCount()
        ]);

        loadingMessage.classList.add("hidden");
        analyticsContent.classList.remove("hidden");

        // Populate Current Tab
        const notOnLda = totalStudents - ldaStudents;
        const percentageOnLda = totalStudents > 0 ? ((ldaStudents / totalStudents) * 100).toFixed(1) : 0;

        document.getElementById("total-students").textContent = totalStudents;
        document.getElementById("lda-students").textContent = ldaStudents;
        document.getElementById("lda-percentage").textContent = `${percentageOnLda}%`;
        renderPieChart(ldaStudents, notOnLda);

        // Populate Projection Tab
        const tomorrowTotal = ldaStudents + projectedStudents;
        document.getElementById("lda-students-proj").textContent = ldaStudents;
        document.getElementById("projected-students").textContent = projectedStudents;
        document.getElementById("tomorrow-total-lda").textContent = tomorrowTotal;

    } catch (error) {
        loadingMessage.classList.add("hidden");
        showError(error.message);
        console.error(error);
    }
}

async function getTotalStudentCount() {
    return await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getItem("Master List");
        const range = sheet.getUsedRange();
        range.load("rowCount");
        await context.sync();
        // Subtract 1 for the header row
        return range.rowCount > 0 ? range.rowCount - 1 : 0;
    });
}

async function getLdaStudentCount() {
    return await Excel.run(async (context) => {
        const worksheets = context.workbook.worksheets;
        worksheets.load("items/name");
        await context.sync();

        let latestLdaSheet = null;
        let latestDate = null;

        worksheets.items.forEach(sheet => {
            if (sheet.name.startsWith("LDA ")) {
                const datePart = sheet.name.substring(4).split(" ")[0]; // Handle names like "LDA 7-31-2025 (2)"
                const date = new Date(datePart.replace(/(\d{1,2})-(\d{1,2})-(\d{4})/, '$3-$1-$2'));
                if (!isNaN(date.getTime())) {
                    if (!latestDate || date > latestDate) {
                        latestDate = date;
                        latestLdaSheet = sheet;
                    }
                }
            }
        });

        if (!latestLdaSheet) {
            throw new Error("No LDA sheet found. Please create an LDA report first.");
        }

        const ldaTable = latestLdaSheet.tables.getItemAt(0);
        const bodyRange = ldaTable.getDataBodyRange();
        bodyRange.load("rowCount");
        await context.sync();
        
        return bodyRange.rowCount;
    });
}

async function getProjectedLdaStudentCount() {
    return await Excel.run(async (context) => {
        const settings = await getSettings();
        const daysOutFilter = settings.createlda.daysOutFilter || 6;
        const projectionThreshold = daysOutFilter - 1;

        const sheet = context.workbook.worksheets.getItem("Master List");
        const range = sheet.getUsedRange();
        range.load("values");
        await context.sync();

        const values = range.values;
        const headers = values[0].map(h => String(h || '').toLowerCase());
        const daysOutColIdx = findColumnIndex(headers, CONSTANTS.COLUMN_MAPPINGS.daysOut);

        if (daysOutColIdx === -1) {
            throw new Error("'Days Out' column not found in Master List.");
        }

        let projectedCount = 0;
        for (let i = 1; i < values.length; i++) {
            const daysOut = values[i][daysOutColIdx];
            if (typeof daysOut === 'number' && daysOut === projectionThreshold) {
                projectedCount++;
            }
        }
        return projectedCount;
    });
}

function renderPieChart(ldaCount, notOnLdaCount) {
    const ctx = document.getElementById('ldaPieChart').getContext('2d');
    new Chart(ctx, {
        type: 'pie',
        data: {
            labels: ['On LDA List', 'Not on LDA List'],
            datasets: [{
                label: 'Student Distribution',
                data: [ldaCount, notOnLdaCount],
                backgroundColor: [
                    'rgba(59, 130, 246, 0.7)', // blue-500
                    'rgba(209, 213, 219, 0.7)' // gray-300
                ],
                borderColor: [
                    'rgba(59, 130, 246, 1)',
                    'rgba(209, 213, 219, 1)'
                ],
                borderWidth: 1
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: {
                    position: 'top',
                },
                tooltip: {
                    callbacks: {
                        label: function(context) {
                            let label = context.label || '';
                            if (label) {
                                label += ': ';
                            }
                            if (context.parsed !== null) {
                                label += context.parsed;
                            }
                            return label;
                        }
                    }
                }
            }
        }
    });
}

function showError(message) {
    const errorContainer = document.getElementById("error-message");
    const errorText = document.getElementById("error-text");
    if (errorContainer && errorText) {
        errorText.textContent = `Error: ${message}`;
        errorContainer.classList.remove("hidden");
    }
}

// --- Helper Functions (from commands/utils.js) ---

function findColumnIndex(headers, possibleNames) {
    for (const name of possibleNames) {
        const index = headers.indexOf(name);
        if (index !== -1) {
            return index;
        }
    }
    return -1;
}

async function getSettings() {
    await new Promise((resolve) => {
        Office.context.document.settings.refreshAsync(asyncResult => {
            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                console.error("Error refreshing settings: " + asyncResult.error.message);
            }
            resolve();
        });
    });

    const settingsString = Office.context.document.settings.get(CONSTANTS.SETTINGS_KEY);
    const defaults = {
        createlda: {
            daysOutFilter: 6,
            includeFailingList: true,
            hideLeftoverColumns: true,
            treatEmptyGradesAsZero: false,
            ldaColumns: ['Assigned', 'StudentName', 'StudentNumber', 'LDA', 'Days Out', 'grade', 'Phone', 'Outreach']
        },
        userProfile: {
            name: Office.context.displayName || "",
            hasSeenWelcomeMessage: false
        }
    };

    if (settingsString) {
        try {
            const settings = JSON.parse(settingsString);
            settings.createlda = { ...defaults.createlda, ...(settings.createlda || {}) };
            settings.userProfile = { ...defaults.userProfile, ...(settings.userProfile || {}) };
            return settings;
        } catch (e) {
            console.error("Error parsing settings, returning defaults:", e);
            return defaults;
        }
    }
    return defaults;
}
