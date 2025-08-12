Office.onReady((info) => {
    if (info.host === Office.HostType.Excel) {
        document.addEventListener("DOMContentLoaded", run);
    }
});

// --- CONFIGURATION ---
// In a real-world scenario, you might fetch this list from a server.
// For this add-in, we'll list the known formula files here.
const FORMULA_FILES = [
    "standard-risk-model.json",
    "re-entry-risk-model.json",
    "engagement-momentum-model.json"
];

const FORMULAS_PATH = "Formulas/";

async function run() {
    const selector = document.getElementById('formula-selector');
    selector.addEventListener('change', () => displayFormula(selector.value));
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
            if (!response.ok) {
                throw new Error(`HTTP error! status: ${response.status}`);
            }
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

    if (!fileName) {
        container.innerHTML = '<p class="text-center text-gray-500">Please select a formula to view its details.</p>';
        return;
    }

    // Hide container, show spinner, and clear previous content
    container.classList.add('hidden');
    spinner.classList.remove('hidden');
    container.innerHTML = ''; 

    try {
        const response = await fetch(`${FORMULAS_PATH}${fileName}`);
        if (!response.ok) throw new Error(`Failed to fetch ${fileName}`);
        const formula = await response.json();
        
        renderFormulaDetails(formula, container);

    } catch (error) {
        container.innerHTML = `<p class="text-center text-red-500">Error loading formula: ${error.message}</p>`;
        console.error(error);
    } finally {
        // Always hide spinner and show container after operation
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
            <p class="text-xs text-gray-500 mt-2">Maximum Possible Score: <strong>${formula.maxScore}</strong></p>
        </div>
    `;

    // Render Components
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

    // Render Modifiers
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
