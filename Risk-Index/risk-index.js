Office.onReady((info) => {
    if (info.host === Office.HostType.Excel) {
        document.addEventListener("DOMContentLoaded", run);
    }
});

// --- CONFIGURATION ---
const FORMULA_FILES = [
    "standard-risk-model.json",
    "re-entry-risk-model.json",
    "engagement-momentum-model.json"
];
const FORMULAS_PATH = "Formulas/";
const MASTER_LIST_SHEET = "Master List";

// --- STATE ---
let currentFormula = null;

// --- MAPPINGS ---
// Maps formula columns to possible aliases in the Excel sheet
const COLUMN_ALIASES = {
    "Grade": ["grade", "course grade"],
    "Days Out": ["days out", "daysout"],
    "Contact Attempts": ["contact attempts"],
    "Times Contacted": ["times contacted"],
    "Prev. Grade": ["prev. grade"],
    "GPA": ["gpa"],
    "Re-entry": ["re-entry"],
    "Status": ["status"],
    "Grade Trend": ["grade trend"],
    "Course Missing Assignments": ["course missing assignments"],
    "Course Zero Assignments": ["course zero assignments"]
};


async function run() {
    const selector = document.getElementById('formula-selector');
    selector.addEventListener('change', () => displayFormula(selector.value));
    document.getElementById('check-compatibility-button').addEventListener('click', checkCompatibility);
    await loadFormulaOptions();
}

/**
 * Loads the available formula files into the dropdown selector.
 */
async function loadFormulaOptions() {
    const selector = document.getElementById('formula-selector');
    selector.innerHTML = '<option value="">Select a model</option>';

    for (const fileName of FORMULA_FILES) {
        try {
            const response = await fetch(`${FORMULAS_PATH}${fileName}`);
            if (!response.ok) throw new Error(`HTTP error! status: ${response.status}`);
            const formula = await response.json();
            const option = document.createElement('option');
            option.value = fileName;
            option.textContent = formula.modelName || fileName;
            selector.appendChild(option);
        } catch (error) {
            console.error(`Failed to load formula: ${fileName}`, error);
        }
    }

    if (selector.options.length <= 1) {
        selector.innerHTML = '<option>No formulas found</option>';
    }
}

/**
 * Fetches and displays the details of a selected formula JSON file.
 * @param {string} fileName The name of the JSON file to load.
 */
async function displayFormula(fileName) {
    const container = document.getElementById('formula-details-container');
    const spinner = document.getElementById('loading-spinner');
    const compatibilitySection = document.getElementById('compatibility-section');
    const resultsContainer = document.getElementById('compatibility-results-container');

    currentFormula = null;
    resultsContainer.innerHTML = ''; // Clear previous results

    if (!fileName) {
        container.innerHTML = '<p class="text-center text-gray-500">Please select a formula to view its details.</p>';
        compatibilitySection.classList.add('hidden');
        return;
    }

    container.classList.add('hidden');
    spinner.classList.remove('hidden');
    compatibilitySection.classList.add('hidden');
    container.innerHTML = ''; 

    try {
        const response = await fetch(`${FORMULAS_PATH}${fileName}`);
        if (!response.ok) throw new Error(`Failed to fetch ${fileName}`);
        
        currentFormula = await response.json();
        renderFormulaDetails(currentFormula, container);
        compatibilitySection.classList.remove('hidden');

    } catch (error) {
        container.innerHTML = `<p class="text-center text-red-500">Error loading formula: ${error.message}</p>`;
        console.error(error);
    } finally {
        spinner.classList.add('hidden');
        container.classList.remove('hidden');
    }
}

/**
 * Renders the parsed formula object into HTML and appends it to the container.
 * @param {object} formula The parsed JSON formula object.
 * @param {HTMLElement} container The container element to render into.
 */
function renderFormulaDetails(formula, container) {
    let html = `
        <div class="bg-white p-4 rounded-lg shadow">
            <h2 class="text-lg font-bold text-gray-800">${formula.modelName} <span class="text-sm font-normal text-gray-500">v${formula.version}</span></h2>
            <p class="text-sm text-gray-600 mt-1">${formula.description}</p>
            <p class="text-xs text-gray-500 mt-2">Author: <strong>${formula.author || 'Unknown'}</strong></p>
            <p class="text-xs text-gray-500 mt-1">Maximum Possible Score: <strong>${formula.maxScore}</strong></p>
        </div>
    `;

    if (formula.components && formula.components.length > 0) {
        html += '<h3 class="text-md font-semibold text-gray-700 mt-6 mb-2">Risk Components</h3>';
        formula.components.forEach(comp => {
            html += `
                <div class="bg-white p-4 rounded-lg shadow mb-3">
                    <div class="flex justify-between items-center">
                        <h4 class="font-bold text-blue-600">${comp.displayName || comp.name}</h4>
                        <span class="text-lg font-bold text-gray-800">${comp.weight} pts</span>
                    </div>
                    <p class="text-xs text-gray-500 mt-2">${comp.calculation}</p>
                </div>
            `;
        });
    }

    if (formula.postCalculationModifiers && formula.postCalculationModifiers.length > 0) {
        html += '<h3 class="text-md font-semibold text-gray-700 mt-6 mb-2">Score Modifiers</h3>';
        formula.postCalculationModifiers.forEach(mod => {
            html += `
                <div class="bg-white p-4 rounded-lg shadow mb-3 border-l-4 border-orange-400">
                    <h4 class="font-bold text-orange-600">${mod.displayName || mod.name}</h4>
                    <p class="text-xs text-gray-500 mt-2">${mod.calculation}</p>
                </div>
            `;
        });
    }

    container.innerHTML = html;
}

/**
 * Checks if the Master List sheet has the columns required by the current formula.
 */
async function checkCompatibility() {
    if (!currentFormula) {
        alert("Please select a formula first.");
        return;
    }

    const resultsContainer = document.getElementById('compatibility-results-container');
    resultsContainer.innerHTML = '<p class="text-center text-gray-500">Checking...</p>';

    try {
        // 1. Gather required columns from the formula
        const requiredColumns = new Set();
        (currentFormula.components || []).forEach(c => {
            if (c.column) requiredColumns.add(c.column);
            if (c.columns) {
                if (c.columns.numerator) requiredColumns.add(c.columns.numerator);
                if (c.columns.denominator) requiredColumns.add(c.columns.denominator);
            }
            if (c.conditionalSource) {
                 if(c.conditionalSource.conditionColumn) requiredColumns.add(c.conditionalSource.conditionColumn);
                 if(c.conditionalSource.sourceColumn) requiredColumns.add(c.conditionalSource.sourceColumn);
            }
        });
        (currentFormula.postCalculationModifiers || []).forEach(m => {
            if (m.conditionColumn) requiredColumns.add(m.conditionColumn);
        });

        // 2. Get headers from the Master List sheet
        const masterHeaders = await Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getItem(MASTER_LIST_SHEET);
            const headerRange = sheet.getRange("1:1").getUsedRange(true);
            headerRange.load("values");
            await context.sync();
            return headerRange.values[0].map(h => String(h || '').toLowerCase());
        });

        // 3. Compare and generate results
        const results = [];
        let allFound = true;
        requiredColumns.forEach(reqCol => {
            const aliases = COLUMN_ALIASES[reqCol] || [reqCol.toLowerCase()];
            const found = aliases.some(alias => masterHeaders.includes(alias));
            if (!found) allFound = false;
            results.push({ name: reqCol, found });
        });

        renderCompatibilityResults(results, allFound);

    } catch (error) {
        let errorMessage = `Error during compatibility check: ${error.message}`;
        if (error.code === 'ItemNotFound') {
            errorMessage = `Error: The "${MASTER_LIST_SHEET}" sheet could not be found.`;
        }
        resultsContainer.innerHTML = `<p class="text-center text-red-500">${errorMessage}</p>`;
        console.error(error);
    }
}

/**
 * Renders the compatibility check results to the UI.
 * @param {Array<{name: string, found: boolean}>} results The array of check results.
 * @param {boolean} allFound Whether all required columns were found.
 */
function renderCompatibilityResults(results, allFound) {
    const container = document.getElementById('compatibility-results-container');
    
    let overallStatus = allFound
        ? `<div class="p-3 mb-3 bg-green-100 text-green-800 rounded-lg text-sm font-semibold">✔ All required columns found. This model is compatible.</div>`
        : `<div class="p-3 mb-3 bg-red-100 text-red-800 rounded-lg text-sm font-semibold">❌ Missing required columns. This model may not run correctly.</div>`;

    let listHtml = '<ul class="space-y-1 text-sm">';
    results.forEach(res => {
        const icon = res.found
            ? '<span class="text-green-500 font-bold mr-2">✔</span>'
            : '<span class="text-red-500 font-bold mr-2">❌</span>';
        const textColor = res.found ? 'text-gray-700' : 'text-red-700 font-semibold';
        listHtml += `<li class="p-2 bg-gray-50 rounded-md flex items-center ${textColor}">${icon} ${res.name}</li>`;
    });
    listHtml += '</ul>';

    container.innerHTML = overallStatus + listHtml;
}
