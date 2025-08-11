Office.onReady((info) => {
    if (info.host === Office.HostType.Excel) {
        run();
    }
});

const CONSTANTS = {
    SETTINGS_KEY: "studentRetentionSettings",
    COLUMN_MAPPINGS: {
        daysOut: ["days out", "daysout"],
        studentName: ["studentname", "student name"],
        outreach: ["outreach"]
    }
};

let trendsChartInstance = null;
let fullTrendsData = null;
let currentTrendPeriod = 'month';
let allDataLoaded = false;

function logAnalyticsProgress(message) {
    const logContainer = document.getElementById('log-container');
    if (logContainer) {
        const p = document.createElement('p');
        p.textContent = `> ${message}`;
        logContainer.appendChild(p);
        logContainer.scrollTop = logContainer.scrollHeight;
    }
}

function updateProgressBar(processed, total, statusText) {
    const progressBarFill = document.getElementById('progress-bar-fill');
    const statusTextElement = document.getElementById('progress-status-text');
    
    if (progressBarFill && statusTextElement) {
        const percentage = total > 0 ? (processed / total) * 100 : 0;
        progressBarFill.style.width = `${percentage}%`;
        
        let text = statusText;
        if (total > 0) {
            text += ` (${processed}/${total})`;
        }
        statusTextElement.textContent = text;
    }
}

async function run() {
    // Reset state variables on each run to prevent issues on re-opening the pane
    if (trendsChartInstance) {
        trendsChartInstance.destroy();
    }
    trendsChartInstance = null;
    fullTrendsData = null;
    allDataLoaded = false;
    currentTrendPeriod = 'month'; // Reset to default

    setupTabs();
    setupTrendFilters();
    setupLogToggle();
    await setupDaysOutFilter();
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
                tabs.forEach(item => {
                    if (item.button) {
                        item.button.classList.remove('active', 'border-blue-500', 'text-blue-600');
                        item.button.classList.add('border-transparent', 'text-gray-500', 'hover:text-gray-700', 'hover:border-gray-300');
                    }
                    if (item.panel) item.panel.classList.add('hidden');
                });
                if (tab.button) {
                    tab.button.classList.add('active', 'border-blue-500', 'text-blue-600');
                    tab.button.classList.remove('border-transparent', 'text-gray-500');
                }
                if (tab.panel) tab.panel.classList.remove('hidden');
            });
        }
    });
}

function setupTrendFilters() {
    const filters = [
        { id: 'filter-week', period: 'week', text: 'This Week' },
        { id: 'filter-month', period: 'month', text: 'This Month' },
        { id: 'filter-year', period: 'year', text: 'This Year' }
    ];

    filters.forEach(filter => {
        const button = document.getElementById(filter.id);
        if (button) {
            button.addEventListener('click', async () => {
                currentTrendPeriod = filter.period;

                if (filter.period === 'year' && !allDataLoaded) {
                    showChartLoading(true);
                    try {
                        fullTrendsData = await getTrendsData('year', updateProgressBar); // Fetch all data
                        allDataLoaded = true;
                    } catch (error) {
                        showError("Failed to load full year data.");
                        console.error(error);
                        showChartLoading(false);
                        return;
                    }
                }

                applyTrendFilter();
                document.querySelectorAll('.filter-button').forEach(btn => btn.classList.remove('active'));
                button.classList.add('active');
                document.getElementById('median-engagement-range').textContent = `for ${filter.text}`;
            });
        }
    });
}

async function setupDaysOutFilter() {
    const filterInput = document.getElementById('trends-days-out-filter');
    if (filterInput) {
        const settings = await getSettings();
        filterInput.value = settings.createlda.daysOutFilter || 6;
        
        filterInput.addEventListener('change', () => {
            applyTrendFilter();
        });
    }
}

function setupLogToggle() {
    const toggleButton = document.getElementById('toggle-log-button');
    const logContainer = document.getElementById('log-container');

    if (toggleButton && logContainer) {
        toggleButton.addEventListener('click', () => {
            const isHidden = logContainer.classList.toggle('hidden');
            toggleButton.textContent = isHidden ? 'Show Logs' : 'Hide Logs';
        });
    }
}


async function loadAnalytics() {
    const loadingMessage = document.getElementById("loading-message");
    const analyticsContent = document.getElementById("analytics-content");
    
    try {
        updateProgressBar(0, 0, "Fetching data for Current and Projection tabs...");
        const [totalStudents, ldaStudents, projectedStudents, engagementData, monthlyEngagementRate] = await Promise.all([
            getTotalStudentCount(),
            getLdaStudentCount(),
            getProjectedLdaStudentCount(),
            getLdaEngagementData(),
            getMonthlyEngagementRate()
        ]);
        
        const trendsData = await getTrendsData('month', updateProgressBar);
        
        fullTrendsData = trendsData;
        allDataLoaded = false;

        loadingMessage.classList.add("hidden");
        document.getElementById('toggle-log-button').textContent = 'Show Logs'; // Set button text
        analyticsContent.classList.remove("hidden");

        const notOnLda = totalStudents - ldaStudents;
        const percentageOnLda = totalStudents > 0 ? ((ldaStudents / totalStudents) * 100).toFixed(1) : 0;

        document.getElementById("total-students").textContent = totalStudents;
        document.getElementById("lda-students").textContent = ldaStudents;
        document.getElementById("lda-percentage").textContent = `${percentageOnLda}%`;
        renderPieChart(ldaStudents, notOnLda);
        renderEngagementPieChart(engagementData.engaged, engagementData.notEngaged);

        // Projection Tab Calculations
        const tomorrowTotal = ldaStudents + projectedStudents;
        const projectedEngaged = Math.round(tomorrowTotal * monthlyEngagementRate);
        const projectedNetLda = tomorrowTotal - projectedEngaged;

        document.getElementById("lda-students-proj").textContent = ldaStudents;
        document.getElementById("projected-students").textContent = projectedStudents;
        document.getElementById("tomorrow-total-lda").textContent = tomorrowTotal;
        document.getElementById("monthly-engagement-rate").textContent = `${(monthlyEngagementRate * 100).toFixed(1)}%`;
        document.getElementById("projected-engaged").textContent = projectedEngaged;
        document.getElementById("projected-net-lda").textContent = projectedNetLda;
        
        applyTrendFilter();

    } catch (error) {
        loadingMessage.classList.add("hidden");
        showError(error.message);
        console.error(error);
    }
}

function applyTrendFilter() {
    if (!fullTrendsData) return;
    updateProgressBar(0, 0, "Applying filters and re-rendering trends chart...");

    const period = currentTrendPeriod;
    const daysOutFilter = parseInt(document.getElementById('trends-days-out-filter').value, 10) || 0;

    const now = new Date();
    let startDate;

    switch (period) {
        case 'week':
            startDate = new Date(new Date().setDate(now.getDate() - 7));
            break;
        case 'month':
            startDate = new Date(new Date().setMonth(now.getMonth() - 1));
            break;
        case 'year':
            startDate = new Date(new Date().setFullYear(now.getFullYear() - 1));
            break;
        default:
            startDate = new Date(0);
    }
    
    const chartData = {
        labels: [],
        ldaCounts: [],
        engagedCounts: [],
        sheetNames: []
    };
    
    const sheetsInPeriod = fullTrendsData.filter(sheetData => sheetData.date >= startDate);

    sheetsInPeriod.forEach(sheetData => {
        const filteredStudents = sheetData.students.filter(student => student.daysOut >= daysOutFilter);
        
        const ldaCount = filteredStudents.length;
        const engagedCount = filteredStudents.filter(student => student.isEngaged).length;

        chartData.labels.push(sheetData.label);
        chartData.ldaCounts.push(ldaCount);
        chartData.engagedCounts.push(engagedCount);
        chartData.sheetNames.push(sheetData.sheetName);
    });
    
    const averageLda = chartData.ldaCounts.length > 0 ? (chartData.ldaCounts.reduce((a, b) => a + b, 0) / chartData.ldaCounts.length) : 0;
    chartData.averageLda = averageLda.toFixed(1);

    renderTrendsChart(chartData);
    showChartLoading(false); // Hide loader after rendering
    const medianEngagement = calculateMedian(chartData.engagedCounts);
    document.getElementById("median-engagement").textContent = medianEngagement;
    updateProgressBar(1, 1, "Analytics loaded.");
}


async function getTotalStudentCount() {
    updateProgressBar(0, 0, "Counting total students in Master List...");
    const count = await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getItem("Master List");
        const range = sheet.getUsedRange();
        range.load("rowCount");
        await context.sync();
        return range.rowCount > 0 ? range.rowCount - 1 : 0;
    });
    return count;
}

async function getLatestLdaSheet(context) {
    const worksheets = context.workbook.worksheets;
    worksheets.load("items/name");
    await context.sync();

    let latestLdaSheet = null;
    let latestDate = null;

    worksheets.items.forEach(sheet => {
        if (sheet.name.startsWith("LDA ")) {
            const datePart = sheet.name.substring(4).split(" ")[0];
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
    updateProgressBar(0, 0, "Finding latest LDA sheet and counting students...");
    const count = await Excel.run(async (context) => {
        const latestLdaSheet = await getLatestLdaSheet(context);
        if (!latestLdaSheet) {
            throw new Error("No LDA sheet found. Please create an LDA report first.");
        }
        latestLdaSheet.tables.load("items/name");
        await context.sync();
        if (latestLdaSheet.tables.items.length === 0) return 0;

        const ldaTable = latestLdaSheet.tables.getItemAt(0);
        const bodyRange = ldaTable.getDataBodyRange();
        bodyRange.load("rowCount");
        await context.sync();
        
        return bodyRange.rowCount;
    });
    return count;
}

async function getProjectedLdaStudentCount() {
    updateProgressBar(0, 0, "Calculating projected LDA students for tomorrow...");
    const count = await Excel.run(async (context) => {
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

        if (daysOutColIdx === -1) throw new Error("'Days Out' column not found in Master List.");

        let projectedCount = 0;
        for (let i = 1; i < values.length; i++) {
            const daysOut = values[i][daysOutColIdx];
            if (typeof daysOut === 'number' && daysOut === projectionThreshold) {
                projectedCount++;
            }
        }
        return projectedCount;
    });
    return count;
}

async function getLdaEngagementData() {
    updateProgressBar(0, 0, "Calculating engagement on latest LDA sheet...");
    const data = await Excel.run(async (context) => {
        const latestLdaSheet = await getLatestLdaSheet(context);
        if (!latestLdaSheet) {
            console.warn("No LDA sheet found for engagement analytics.");
            return { engaged: 0, notEngaged: 0 };
        }
        
        latestLdaSheet.tables.load("items/name");
        await context.sync();
        if (latestLdaSheet.tables.items.length === 0) return { engaged: 0, notEngaged: 0 };

        const ldaTable = latestLdaSheet.tables.getItemAt(0);
        const headerRange = ldaTable.getHeaderRowRange();
        const bodyRange = ldaTable.getDataBodyRange();
        headerRange.load("values");
        bodyRange.load("rowCount");
        await context.sync();

        const totalLdaStudents = bodyRange.rowCount;
        if (totalLdaStudents === 0) return { engaged: 0, notEngaged: 0 };

        const headers = headerRange.values[0].map(h => String(h || '').toLowerCase());
        const outreachColIdx = findColumnIndex(headers, CONSTANTS.COLUMN_MAPPINGS.outreach);

        if (outreachColIdx === -1) throw new Error("'Outreach' column not found in the LDA sheet.");

        let engagedCount = 0;
        const greenShades = ["#C6EFCE", "#92D050", "#00B050", "#90EE90"]; 

        for (let i = 0; i < totalLdaStudents; i++) {
            const rowRange = bodyRange.getRow(i);
            const outreachCell = rowRange.getCell(0, outreachColIdx);
            outreachCell.load("format/fill/color");
            await context.sync();

            const cellColor = outreachCell.format.fill.color;
            if (cellColor && greenShades.some(shade => cellColor.toUpperCase().includes(shade.toUpperCase()))) {
                engagedCount++;
            }
        }
        
        return { engaged: engagedCount, notEngaged: totalLdaStudents - engagedCount };
    });
    return data;
}

async function getTrendsData(period = 'year', progressCallback) {
    const callback = progressCallback || (() => {}); // Use provided callback or a no-op
    return await Excel.run(async (context) => {
        callback(0, 0, "Searching for all LDA sheets in the workbook...");
        const worksheets = context.workbook.worksheets;
        worksheets.load("items/name");
        await context.sync();

        let ldaSheets = [];
        worksheets.items.forEach(sheet => {
            if (sheet.name.startsWith("LDA ")) {
                const datePart = sheet.name.substring(4).split(" ")[0];
                const date = new Date(datePart.replace(/(\d{1,2})-(\d{1,2})-(\d{4})/, '$3-$1-$2'));
                if (!isNaN(date.getTime())) {
                    ldaSheets.push({ sheet, date });
                }
            }
        });
        
        if (period === 'month') {
            const oneMonthAgo = new Date();
            oneMonthAgo.setMonth(oneMonthAgo.getMonth() - 1);
            ldaSheets = ldaSheets.filter(s => s.date >= oneMonthAgo);
        }

        ldaSheets.sort((a, b) => a.date - b.date);
        const totalSheets = ldaSheets.length;
        callback(0, totalSheets, `Found ${totalSheets} LDA sheets to process.`);

        const trendsData = [];
        let processedSheets = 0;

        for (const { sheet, date } of ldaSheets) {
            callback(processedSheets, totalSheets, `Processing sheet: ${sheet.name}...`);
            const sheetData = {
                label: date.toLocaleDateString(),
                date: date,
                sheetName: sheet.name,
                students: []
            };

            sheet.tables.load("items/name");
            await context.sync();

            if (sheet.tables.items.length === 0) {
                trendsData.push(sheetData);
                processedSheets++;
                continue;
            }

            const ldaTable = sheet.tables.getItemAt(0);
            const headerRange = ldaTable.getHeaderRowRange();
            const bodyRange = ldaTable.getDataBodyRange();
            headerRange.load("values");
            bodyRange.load("values, rowCount");
            await context.sync();

            const headers = headerRange.values[0].map(h => String(h || '').toLowerCase());
            const outreachColIdx = findColumnIndex(headers, CONSTANTS.COLUMN_MAPPINGS.outreach);
            const daysOutColIdx = findColumnIndex(headers, CONSTANTS.COLUMN_MAPPINGS.daysOut);

            if (outreachColIdx === -1 || daysOutColIdx === -1) {
                trendsData.push(sheetData);
                processedSheets++;
                continue;
            }

            const greenShades = ["#C6EFCE", "#92D050", "#00B050", "#90EE90"];
            const values = bodyRange.values;

            for (let i = 0; i < values.length; i++) {
                const rowRange = bodyRange.getRow(i);
                const outreachCell = rowRange.getCell(0, outreachColIdx);
                outreachCell.load("format/fill/color");
                await context.sync();
                
                const cellColor = outreachCell.format.fill.color;
                const isEngaged = cellColor && greenShades.some(shade => cellColor.toUpperCase().includes(shade.toUpperCase()));
                const daysOut = values[i][daysOutColIdx];
                
                sheetData.students.push({
                    daysOut: typeof daysOut === 'number' ? daysOut : 0,
                    isEngaged: isEngaged
                });
            }
            trendsData.push(sheetData);
            processedSheets++;
        }
        callback(processedSheets, totalSheets, "Finished processing sheets.");
        return trendsData;
    });
}

async function getMonthlyEngagementRate() {
    updateProgressBar(0, 0, "Calculating historical monthly engagement rate...");
    const rate = await Excel.run(async (context) => {
        const worksheets = context.workbook.worksheets;
        worksheets.load("items/name");
        await context.sync();

        let ldaSheets = [];
        const oneMonthAgo = new Date();
        oneMonthAgo.setMonth(oneMonthAgo.getMonth() - 1);

        worksheets.items.forEach(sheet => {
            if (sheet.name.startsWith("LDA ")) {
                const datePart = sheet.name.substring(4).split(" ")[0];
                const date = new Date(datePart.replace(/(\d{1,2})-(\d{1,2})-(\d{4})/, '$3-$1-$2'));
                if (!isNaN(date.getTime()) && date >= oneMonthAgo) {
                    ldaSheets.push(sheet);
                }
            }
        });

        if (ldaSheets.length === 0) {
            return 0;
        }

        let totalStudentsInMonth = 0;
        let totalEngagedInMonth = 0;
        const greenShades = ["#C6EFCE", "#92D050", "#00B050", "#90EE90"];

        for (const sheet of ldaSheets) {
            sheet.tables.load("items/name");
            await context.sync();
            if (sheet.tables.items.length === 0) continue;

            const table = sheet.tables.getItemAt(0);
            const headerRange = table.getHeaderRowRange();
            const bodyRange = table.getDataBodyRange();
            headerRange.load("values");
            bodyRange.load("rowCount");
            await context.sync();

            const headers = headerRange.values[0].map(h => String(h || '').toLowerCase());
            const outreachColIdx = findColumnIndex(headers, CONSTANTS.COLUMN_MAPPINGS.outreach);
            if (outreachColIdx === -1) continue;

            const rowCount = bodyRange.rowCount;
            totalStudentsInMonth += rowCount;

            for (let i = 0; i < rowCount; i++) {
                const cell = bodyRange.getRow(i).getCell(0, outreachColIdx);
                cell.load("format/fill/color");
                await context.sync();
                const cellColor = cell.format.fill.color;
                if (cellColor && greenShades.some(shade => cellColor.toUpperCase().includes(shade.toUpperCase()))) {
                    totalEngagedInMonth++;
                }
            }
        }

        if (totalStudentsInMonth === 0) return 0;
        return totalEngagedInMonth / totalStudentsInMonth;
    });
    return rate;
}

function calculateMedian(numbers) {
    if (!numbers || numbers.length === 0) return 0;
    const sorted = [...numbers].sort((a, b) => a - b);
    const middleIndex = Math.floor(sorted.length / 2);
    if (sorted.length % 2 === 0) {
        return ((sorted[middleIndex - 1] + sorted[middleIndex]) / 2).toFixed(1);
    } else {
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
                backgroundColor: ['rgba(59, 130, 246, 0.7)', 'rgba(209, 213, 219, 0.7)'],
                borderColor: ['rgba(59, 130, 246, 1)', 'rgba(209, 213, 219, 1)'],
                borderWidth: 1
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: { position: 'top' },
                tooltip: {
                    callbacks: {
                        label: function(context) {
                            let label = context.label || '';
                            if (label) {
                                label += ': ';
                            }
                            const value = context.raw;
                            const sum = context.chart.data.datasets[0].data.reduce((a, b) => a + b, 0);
                            const percentage = sum > 0 ? (value / sum * 100).toFixed(1) : 0;
                            label += `${percentage}%`;
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
                backgroundColor: ['rgba(34, 197, 94, 0.7)', 'rgba(209, 213, 219, 0.7)'],
                borderColor: ['rgba(34, 197, 94, 1)', 'rgba(209, 213, 219, 1)'],
                borderWidth: 1
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: { position: 'top' },
                tooltip: {
                    callbacks: {
                        label: function(context) {
                            let label = context.label || '';
                            if (label) {
                                label += ': ';
                            }
                            const value = context.raw;
                            const sum = context.chart.data.datasets[0].data.reduce((a, b) => a + b, 0);
                            const percentage = sum > 0 ? (value / sum * 100).toFixed(1) : 0;
                            label += `${percentage}%`;
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
    if (trendsChartInstance) {
        trendsChartInstance.destroy();
    }

    const averageData = new Array(trendsData.labels.length).fill(trendsData.averageLda);

    trendsChartInstance = new Chart(ctx, {
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
            }, {
                label: `Average LDA (${trendsData.averageLda})`,
                data: averageData,
                borderColor: 'rgba(239, 68, 68, 0.8)',
                borderDash: [5, 5],
                fill: false,
                tension: 0.1,
                pointRadius: 0
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            scales: { y: { beginAtZero: true } },
            onClick: (event) => {
                const points = trendsChartInstance.getElementsAtEventForMode(event, 'nearest', { intersect: true }, true);
                if (points.length) {
                    const firstPoint = points[0];
                    const index = firstPoint.index;
                    const sheetName = trendsData.sheetNames[index];
                    if (sheetName) {
                        navigateToSheet(sheetName);
                    }
                }
            }
        }
    });
}

/**
 * Navigates to a specific worksheet by its name.
 * @param {string} sheetName The name of the sheet to activate.
 */
async function navigateToSheet(sheetName) {
    try {
        await Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getItem(sheetName);
            sheet.load("visibility");
            await context.sync();
            
            if (sheet.visibility !== Excel.SheetVisibility.visible) {
                sheet.visibility = Excel.SheetVisibility.visible;
            }
            
            sheet.activate();
            await context.sync();
        });
    } catch (error) {
        console.error(`Error navigating to sheet "${sheetName}":`, error);
        showError(`Could not navigate to sheet: ${sheetName}`);
    }
}


function showError(message) {
    const errorContainer = document.getElementById("error-message");
    const errorText = document.getElementById("error-text");
    if (errorContainer && errorText) {
        errorText.textContent = `Error: ${message}`;
        errorContainer.classList.remove("hidden");
    }
}

function findColumnIndex(headers, possibleNames) {
    for (const name of possibleNames) {
        const index = headers.indexOf(name);
        if (index !== -1) return index;
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
        userProfile: { name: Office.context.displayName || "", hasSeenWelcomeMessage: false }
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

function showChartLoading(isLoading) {
    const canvas = document.getElementById('trendsChart');
    const container = canvas.parentElement;
    
    let loader = container.querySelector('.chart-loader');
    if (isLoading) {
        if (!loader) {
            loader = document.createElement('div');
            loader.className = 'chart-loader absolute inset-0 bg-white bg-opacity-75 flex items-center justify-center';
            loader.innerHTML = `<p class="text-gray-600">Loading full year data...</p>`;
            container.appendChild(loader);
        }
        loader.style.display = 'flex';
    } else {
        if (loader) {
            loader.style.display = 'none';
        }
    }
}
