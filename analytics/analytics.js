Office.onReady((info) => {
    if (info.host === Office.HostType.Excel) {
        document.addEventListener("DOMContentLoaded", run);
    }
});

async function run() {
    try {
        await loadAnalytics();
    } catch (error) {
        showError(error.message);
        console.error(error);
    }
}

async function loadAnalytics() {
    const loadingMessage = document.getElementById("loading-message");
    const analyticsContent = document.getElementById("analytics-content");
    
    try {
        const [totalStudents, ldaStudents] = await Promise.all([
            getTotalStudentCount(),
            getLdaStudentCount()
        ]);

        loadingMessage.classList.add("hidden");
        analyticsContent.classList.remove("hidden");

        const notOnLda = totalStudents - ldaStudents;
        const percentageOnLda = totalStudents > 0 ? ((ldaStudents / totalStudents) * 100).toFixed(1) : 0;

        document.getElementById("total-students").textContent = totalStudents;
        document.getElementById("lda-students").textContent = ldaStudents;
        document.getElementById("lda-percentage").textContent = `${percentageOnLda}%`;

        renderPieChart(ldaStudents, notOnLda);

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
                const datePart = sheet.name.substring(4);
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
