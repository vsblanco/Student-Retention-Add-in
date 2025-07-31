Office.onReady((info) => {
    if (info.host === Office.HostType.Excel) {
        document.addEventListener("DOMContentLoaded", run);
    }
});

const CONSTANTS = {
    SETTINGS_KEY: "studentRetentionSettings",
    COLUMN_MAPPINGS: {
        daysOut: ["days out", "daysout"],
        studentName: ["studentname", "student name"],
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
    const tabTrends = document.getElementById('tab-trends');
    const panelCurrent = document.getElementById('panel-current');
    const panelProjection = document.getElementById('panel-projection');
    const panelTrends = document.getElementById('panel-trends');

    const tabs = [
        { button: tabCurrent, panel: panelCurrent },
        { button: tabProjection, panel: panelProjection },
        { button: tabTrends, panel: panelTrends }
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
        const [totalStudents, ldaStudents, projectedStudents, engagementData, trendsData] = await Promise.all([
            getTotalStudentCount(),
            getLdaStudentCount(),
            getProjectedLdaStudentCount(),
            getLdaEngagementData(),
            getTrendsData()
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
        renderEngagementPieChart(engagementData.engaged, engagementData.notEngaged);

        // Populate Projection Tab
        const tomorrowTotal = ldaStudents + projectedStudents;
        document.getElementById("lda-students-proj").textContent = ldaStudents;
        document.getElementById("projected-students").textContent = projectedStudents;
        document.getElementById("tomorrow-total-lda").textContent = tomorrowTotal;
        
        // Populate Trends Tab
        renderTrendsChart(trendsData);
        const medianEngagement = calculateMedian(trendsData.engagedCounts);
        document.getElementById("median-engagement").textContent = medianEngagement;

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

async function getLatestLdaSheet(context) {
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

    return latestLdaSheet;
}


async function getLdaStudentCount() {
    return await Excel.run(async (context) => {
        const latestLdaSheet = await getLatestLdaSheet(context);
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

async function getLdaEngagementData() {
    return await Excel.run(async (context) => {
        const latestLdaSheet = await getLatestLdaSheet(context);
        if (!latestLdaSheet) {
            console.warn("No LDA sheet found for engagement analytics.");
            return { engaged: 0, notEngaged: 0 };
        }

        const ldaTable = latestLdaSheet.tables.getItemAt(0);
        const headerRange = ldaTable.getHeaderRowRange();
        const bodyRange = ldaTable.getDataBodyRange();
        headerRange.load("values");
        bodyRange.load("rowCount, format/fill/color"); // Load colors for the whole body
        await context.sync();

        const totalLdaStudents = bodyRange.rowCount;
        if (totalLdaStudents === 0) {
            return { engaged: 0, notEngaged: 0 };
        }

        const headers = headerRange.values[0].map(h => String(h || '').toLowerCase());
        const studentNameColIdx = findColumnIndex(headers, CONSTANTS.COLUMN_MAPPINGS.studentName);

        if (studentNameColIdx === -1) {
            throw new Error("'StudentName' column not found in the LDA sheet.");
        }

        let engagedCount = 0;
        const greenShades = ["#C6EFCE", "#92D050", "#00B050", "#90EE90"]; 

        const colors = bodyRange.format.fill.color;
        if (colors) {
            for (let i = 0; i < totalLdaStudents; i++) {
                // FIX: Check if the row of colors exists before accessing a cell
                if (colors[i]) {
                    const cellColor = colors[i][studentNameColIdx];
                    if (cellColor && greenShades.some(shade => cellColor.toUpperCase().includes(shade.toUpperCase()))) {
                        engagedCount++;
                    }
                }
            }
        }
        
        return {
            engaged: engagedCount,
            notEngaged: totalLdaStudents - engagedCount
        };
    });
}

async function getTrendsData() {
    return await Excel.run(async (context) => {
        const worksheets = context.workbook.worksheets;
        worksheets.load("items/name");
        await context.sync();

        const ldaSheets = [];
        worksheets.items.forEach(sheet => {
            if (sheet.name.startsWith("LDA ")) {
                const datePart = sheet.name.substring(4).split(" ")[0];
                const date = new Date(datePart.replace(/(\d{1,2})-(\d{1,2})-(\d{4})/, '$3-$1-$2'));
                if (!isNaN(date.getTime())) {
                    ldaSheets.push({ sheet, date });
                }
            }
        });

        ldaSheets.sort((a, b) => a.date - b.date);

        const trendsData = {
            labels: [],
            ldaCounts: [],
            engagedCounts: []
        };

        for (const { sheet, date } of ldaSheets) {
            trendsData.labels.push(date.toLocaleDateString());

            // FIX: Check if the sheet has a table before proceeding
            sheet.tables.load("items/name");
            await context.sync();

            if (sheet.tables.items.length === 0) {
                trendsData.ldaCounts.push(0);
                trendsData.engagedCounts.push(0);
                continue; // Skip to the next sheet
            }

            const ldaTable = sheet.tables.getItemAt(0);
            const headerRange = ldaTable.getHeaderRowRange();
            const bodyRange = ldaTable.getDataBodyRange();
            headerRange.load("values");
            bodyRange.load("rowCount, format/fill/color");
            await context.sync();

            const totalLdaStudents = bodyRange.rowCount;
            trendsData.ldaCounts.push(totalLdaStudents);

            const headers = headerRange.values[0].map(h => String(h || '').toLowerCase());
            const studentNameColIdx = findColumnIndex(headers, CONSTANTS.COLUMN_MAPPINGS.studentName);

            if (studentNameColIdx === -1) {
                trendsData.engagedCounts.push(0);
                continue;
            }

            let engagedCount = 0;
            const greenShades = ["#C6EFCE", "#92D050", "#00B050", "#90EE90"];
            const colors = bodyRange.format.fill.color;
            if (colors) {
                for (let i = 0; i < totalLdaStudents; i++) {
                    if (colors[i]) {
                        const cellColor = colors[i][studentNameColIdx];
                        if (cellColor && greenShades.some(shade => cellColor.toUpperCase().includes(shade.toUpperCase()))) {
                            engagedCount++;
                        }
                    }
                }
            }
            trendsData.engagedCounts.push(engagedCount);
        }

        return trendsData;
    });
}

function calculateMedian(numbers) {
    if (!numbers || numbers.length === 0) {
        return 0;
    }

    const sorted = [...numbers].sort((a, b) => a - b);
    const middleIndex = Math.floor(sorted.length / 2);

    if (sorted.length % 2 === 0) {
        // Even number of elements
        return ((sorted[middleIndex - 1] + sorted[middleIndex]) / 2).toFixed(1);
    } else {
        // Odd number of elements
        return sorted[middleIndex];
    }
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

function renderEngagementPieChart(engagedCount, notEngagedCount) {
    const ctx = document.getElementById('engagementPieChart').getContext('2d');
    new Chart(ctx, {
        type: 'pie',
        data: {
            labels: ['Engaged', 'Not Engaged'],
            datasets: [{
                label: 'Engagement Status',
                data: [engagedCount, notEngagedCount],
                backgroundColor: [
                    'rgba(34, 197, 94, 0.7)', // green-500
                    'rgba(239, 68, 68, 0.7)'  // red-500
                ],
                borderColor: [
                    'rgba(34, 197, 94, 1)',
                    'rgba(239, 68, 68, 1)'
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

function renderTrendsChart(trendsData) {
    const ctx = document.getElementById('trendsChart').getContext('2d');
    new Chart(ctx, {
        type: 'line',
        data: {
            labels: trendsData.labels,
            datasets: [{
                label: 'Total on LDA',
                data: trendsData.ldaCounts,
                borderColor: 'rgba(59, 130, 246, 1)',
                backgroundColor: 'rgba(59, 130, 246, 0.2)',
                fill: true,
                tension: 0.1
            }, {
                label: 'Engaged Students',
                data: trendsData.engagedCounts,
                borderColor: 'rgba(34, 197, 94, 1)',
                backgroundColor: 'rgba(34, 197, 94, 0.2)',
                fill: true,
                tension: 0.1
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            scales: {
                y: {
                    beginAtZero: true
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
