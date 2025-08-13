Office.onReady((info) => {
    if (info.host === Office.HostType.Excel) {
        document.addEventListener("DOMContentLoaded", run);
    }
});

// --- STATE ---
let currentFormula = null;
let componentIdCounter = 0;
let modifierIdCounter = 0;

// --- CONSTANTS ---
const FORMULA_FILES = [
    "standard-risk-model.json",
    "re-entry-risk-model.json",
    "engagement-momentum-model.json"
];
const FORMULAS_PATH = "Formulas/";
const MASTER_LIST_SHEET = "Master List";
const COMPONENT_TYPES = ["value-map", "linear", "linear-multiplier", "inverse-linear", "inverse-ratio"];
const MODIFIER_OPERATORS = ["+", "-", "*", "/"];
const COLUMN_ALIASES = {
    "Grade": ["grade", "course grade"],
    "Days Out": ["days out", "daysout"],
    "Contact Attempts": ["contact attempts"],
    "Times Contacted": ["times contacted"],
    "Prev. Grade": ["prev. grade"],
    "GPA": ["gpa"],
    "Re-entry": ["re-entry", "ReEntry"],
    "Status": ["status"],
    "Grade Trend": ["grade trend"],
    "Course Missing Assignments": ["course missing assignments"],
    "Course Zero Assignments": ["course zero assignments"]
};

/**
 * Main function to initialize the task pane.
 */
async function run() {
    // Viewer listeners
    document.getElementById('formula-selector').addEventListener('change', () => displayFormula(document.getElementById('formula-selector').value));
    document.getElementById('check-compatibility-button').addEventListener('click', checkCompatibility);
    
    // Builder listeners
    document.getElementById('add-component-btn').addEventListener('click', addComponent);
    document.getElementById('add-modifier-btn').addEventListener('click', addModifier);
    document.getElementById('generate-json-btn').addEventListener('click', generateAndCopyJson);

    // View toggler
    document.getElementById('toggle-view-btn').addEventListener('click', toggleView);

    await loadFormulaOptions();
}

/**
 * Toggles between the Viewer and Builder modes.
 */
function toggleView() {
    const viewerMode = document.getElementById('viewer-mode');
    const builderMode = document.getElementById('builder-mode');
    const toggleBtn = document.getElementById('toggle-view-btn');
    const mainTitle = document.getElementById('main-title');

    if (viewerMode.classList.contains('hidden')) {
        // Switch to Viewer
        viewerMode.classList.remove('hidden');
        builderMode.classList.add('hidden');
        toggleBtn.textContent = 'Formula Builder';
        mainTitle.textContent = 'Risk Index Models';
    } else {
        // Switch to Builder
        viewerMode.classList.add('hidden');
        builderMode.classList.remove('hidden');
        toggleBtn.textContent = 'Formula Viewer';
        mainTitle.textContent = 'Formula Builder';
    }
}


// ###################################
// ### FORMULA VIEWER FUNCTIONS    ###
// ###################################

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

async function displayFormula(fileName) {
    const container = document.getElementById('formula-details-container');
    const spinner = document.getElementById('loading-spinner');
    const compatibilitySection = document.getElementById('compatibility-section');
    const resultsContainer = document.getElementById('compatibility-results-container');
    currentFormula = null;
    resultsContainer.innerHTML = '';
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

function renderFormulaDetails(formula, container) {
    let html = `
        <div class="bg-white p-4 rounded-lg shadow">
            <h2 class="text-lg font-bold text-gray-800">${formula.modelName} <span class="text-sm font-normal text-gray-500">v${formula.version}</span></h2>
            <p class="text-sm text-gray-600 mt-1">${formula.description}</p>
            <p class="text-xs text-gray-500 mt-2">Author: <strong>${formula.author || 'Unknown'}</strong></p>
            <p class="text-xs text-gray-500 mt-1">Maximum Possible Score: <strong>${formula.maxScore}</strong></p>
        </div>`;
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
                </div>`;
        });
    }
    if (formula.postCalculationModifiers && formula.postCalculationModifiers.length > 0) {
        html += '<h3 class="text-md font-semibold text-gray-700 mt-6 mb-2">Score Modifiers</h3>';
        formula.postCalculationModifiers.forEach(mod => {
            html += `
                <div class="bg-white p-4 rounded-lg shadow mb-3 border-l-4 border-orange-400">
                    <h4 class="font-bold text-orange-600">${mod.displayName || mod.name}</h4>
                    <p class="text-xs text-gray-500 mt-2">${mod.calculation}</p>
                </div>`;
        });
    }
    container.innerHTML = html;
}

async function checkCompatibility() {
    if (!currentFormula) return;
    const resultsContainer = document.getElementById('compatibility-results-container');
    resultsContainer.innerHTML = '<p class="text-center text-gray-500">Checking...</p>';
    try {
        const requiredColumns = new Set();
        (currentFormula.components || []).forEach(c => {
            if (c.column) requiredColumns.add(c.column);
            if (c.columns) {
                if (c.columns.numerator) requiredColumns.add(c.columns.numerator);
                if (c.columns.denominator) requiredColumns.add(c.columns.denominator);
            }
            if (c.conditionalSource) {
                if (c.conditionalSource.conditionColumn) requiredColumns.add(c.conditionalSource.conditionColumn);
                if (c.conditionalSource.sourceColumn) requiredColumns.add(c.conditionalSource.sourceColumn);
            }
        });
        (currentFormula.postCalculationModifiers || []).forEach(m => {
            if (m.conditionColumn) requiredColumns.add(m.conditionColumn);
        });
        const masterHeaders = await Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getItem(MASTER_LIST_SHEET);
            const headerRange = sheet.getRange("1:1").getUsedRange(true);
            headerRange.load("values");
            await context.sync();
            return headerRange.values[0].map(h => String(h || '').toLowerCase());
        });
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

function renderCompatibilityResults(results, allFound) {
    const container = document.getElementById('compatibility-results-container');
    let overallStatus = allFound
        ? `<div class="p-3 mb-3 bg-green-100 text-green-800 rounded-lg text-sm font-semibold">✔ All required columns found. This model is compatible.</div>`
        : `<div class="p-3 mb-3 bg-red-100 text-red-800 rounded-lg text-sm font-semibold">❌ Missing required columns. This model may not run correctly.</div>`;
    let listHtml = '<ul class="space-y-1 text-sm">';
    results.forEach(res => {
        const icon = res.found ? '<span class="text-green-500 font-bold mr-2">✔</span>' : '<span class="text-red-500 font-bold mr-2">❌</span>';
        const textColor = res.found ? 'text-gray-700' : 'text-red-700 font-semibold';
        listHtml += `<li class="p-2 bg-gray-50 rounded-md flex items-center ${textColor}">${icon} ${res.name}</li>`;
    });
    listHtml += '</ul>';
    container.innerHTML = overallStatus + listHtml;
}


// ####################################
// ### FORMULA BUILDER FUNCTIONS    ###
// ####################################

function addComponent() {
    const id = `comp-${componentIdCounter++}`;
    const container = document.getElementById('components-container');
    const componentCard = document.createElement('div');
    componentCard.id = id;
    componentCard.className = 'component-card border border-gray-200 p-4 rounded-md space-y-3';
    componentCard.innerHTML = `
        <div class="flex justify-between items-center">
            <h3 class="font-semibold text-gray-600">Component #${componentIdCounter}</h3>
            <button class="remove-btn text-red-500 hover:text-red-700 font-bold">Remove</button>
        </div>
        <div class="grid grid-cols-1 md:grid-cols-2 gap-3 text-sm">
            <div><label class="block font-medium">Name</label><input type="text" data-key="name" class="w-full p-1 border rounded" placeholder="e.g., Current Grade"></div>
            <div><label class="block font-medium">Display Name</label><input type="text" data-key="displayName" class="w-full p-1 border rounded" placeholder="e.g., Current Grade"></div>
            <div><label class="block font-medium">Weight (Points)</label><input type="number" data-key="weight" class="w-full p-1 border rounded" value="10"></div>
            <div><label class="block font-medium">Calculation Type</label><select data-key="type" class="w-full p-1 border rounded component-type-select">${COMPONENT_TYPES.map(t => `<option value="${t}">${t}</option>`).join('')}</select></div>
            <div class="md:col-span-2"><label class="block font-medium">Calculation Logic</label><input type="text" data-key="calculation" class="w-full p-1 border rounded" placeholder="e.g., Points assigned based on letter grade."></div>
        </div>
        <div class="type-specific-config mt-3"></div>`;
    container.appendChild(componentCard);
    componentCard.querySelector('.remove-btn').addEventListener('click', () => componentCard.remove());
    const typeSelect = componentCard.querySelector('.component-type-select');
    typeSelect.addEventListener('change', handleTypeChange);
    handleTypeChange({ target: typeSelect });
}

function handleTypeChange(event) {
    const selectElement = event.target;
    const componentCard = selectElement.closest('.component-card');
    const configContainer = componentCard.querySelector('.type-specific-config');
    const selectedType = selectElement.value;
    let configHtml = '<div class="grid grid-cols-1 md:grid-cols-2 gap-3 text-sm">';
    if (["linear", "linear-multiplier", "inverse-linear", "value-map"].includes(selectedType)) {
        configHtml += `<div><label class="block font-medium">Source Column</label><input type="text" data-key="column" class="w-full p-1 border rounded" placeholder="e.g., Grade"></div>`;
    }
    switch (selectedType) {
        case 'linear-multiplier':
            configHtml += `<div><label class="block font-medium">Multiplier</label><input type="number" data-key="multiplier" class="w-full p-1 border rounded" value="1"></div>`;
            break;
        case 'inverse-ratio':
            configHtml += `<div><label class="block font-medium">Numerator Column</label><input type="text" data-key="columns.numerator" class="w-full p-1 border rounded" placeholder="e.g., Contact Attempts"></div><div><label class="block font-medium">Denominator Column</label><input type="text" data-key="columns.denominator" class="w-full p-1 border rounded" placeholder="e.g., Times Contacted"></div>`;
            break;
        case 'value-map':
            configHtml += `<div class="md:col-span-2"><label class="block font-medium">Value Mappings</label><div class="value-map-pairs space-y-2 mt-1"></div><button class="add-map-pair-btn mt-2 text-xs px-2 py-1 bg-gray-200 rounded hover:bg-gray-300">+ Add Pair</button></div>`;
            break;
    }
    configHtml += '</div>';
    configContainer.innerHTML = configHtml;
    if (selectedType === 'value-map') {
        componentCard.querySelector('.add-map-pair-btn').addEventListener('click', (e) => {
            e.preventDefault();
            addValueMapPair(e.target.previousElementSibling);
        });
        addValueMapPair(componentCard.querySelector('.value-map-pairs'));
    }
}

function addValueMapPair(container) {
    const pairDiv = document.createElement('div');
    pairDiv.className = 'flex items-center gap-2';
    pairDiv.innerHTML = `<input type="text" class="w-full p-1 border rounded map-key" placeholder="Key (e.g., A)"><span class="font-bold">:</span><input type="number" class="w-full p-1 border rounded map-value" placeholder="Value (e.g., 0)"><button class="remove-pair-btn text-red-500 hover:text-red-700">&times;</button>`;
    pairDiv.querySelector('.remove-pair-btn').addEventListener('click', () => pairDiv.remove());
    container.appendChild(pairDiv);
}

function addModifier() {
    const id = `mod-${modifierIdCounter++}`;
    const container = document.getElementById('modifiers-container');
    const modifierCard = document.createElement('div');
    modifierCard.id = id;
    modifierCard.className = 'modifier-card border border-gray-200 p-4 rounded-md space-y-3';
    modifierCard.innerHTML = `
        <div class="flex justify-between items-center">
            <h3 class="font-semibold text-gray-600">Modifier #${modifierIdCounter}</h3>
            <button class="remove-btn text-red-500 hover:text-red-700 font-bold">Remove</button>
        </div>
        <div class="grid grid-cols-1 md:grid-cols-2 gap-3 text-sm">
            <div><label class="block font-medium">Name</label><input type="text" data-key="name" class="w-full p-1 border rounded" placeholder="e.g., Re-entry Status"></div>
            <div><label class="block font-medium">Display Name</label><input type="text" data-key="displayName" class="w-full p-1 border rounded" placeholder="e.g., Re-entry Modifier"></div>
            <div><label class="block font-medium">Condition Column</label><input type="text" data-key="conditionColumn" class="w-full p-1 border rounded" placeholder="e.g., Re-entry"></div>
            <div><label class="block font-medium">Condition Value</label><input type="text" data-key="conditionValue" class="w-full p-1 border rounded" placeholder="e.g., true"></div>
            <div class="md:col-span-2"><label class="block font-medium">Calculation</label><input type="text" data-key="calculation" class="w-full p-1 border rounded" placeholder="e.g., If Re-entry, score / 2 + 50"></div>
            <div class="md:col-span-2"><label class="block font-medium">Operations</label><div class="operations-list space-y-2 mt-1"></div><button class="add-operation-btn mt-2 text-xs px-2 py-1 bg-gray-200 rounded hover:bg-gray-300">+ Add Operation</button></div>
        </div>`;
    container.appendChild(modifierCard);
    modifierCard.querySelector('.remove-btn').addEventListener('click', () => modifierCard.remove());
    modifierCard.querySelector('.add-operation-btn').addEventListener('click', (e) => {
        e.preventDefault();
        addOperation(e.target.previousElementSibling);
    });
    addOperation(modifierCard.querySelector('.operations-list'));
}

function addOperation(container) {
    const opDiv = document.createElement('div');
    opDiv.className = 'flex items-center gap-2';
    opDiv.innerHTML = `<select class="w-1/3 p-1 border rounded op-operator">${MODIFIER_OPERATORS.map(op => `<option value="${op}">${op}</option>`).join('')}</select><input type="number" class="w-2/3 p-1 border rounded op-value" placeholder="Value"><button class="remove-op-btn text-red-500 hover:text-red-700">&times;</button>`;
    opDiv.querySelector('.remove-op-btn').addEventListener('click', () => opDiv.remove());
    container.appendChild(opDiv);
}

function generateAndCopyJson() {
    const formula = {
        modelName: document.getElementById('modelName').value,
        author: document.getElementById('author').value,
        description: document.getElementById('description').value,
        version: document.getElementById('version').value,
        maxScore: parseInt(document.getElementById('maxScore').value, 10),
        components: [],
        postCalculationModifiers: []
    };
    document.querySelectorAll('.component-card').forEach(card => {
        const component = {};
        card.querySelectorAll('input[data-key], select[data-key]').forEach(input => {
            let key = input.dataset.key;
            let value = input.type === 'number' ? parseFloat(input.value) : input.value;
            if (key.includes('.')) {
                const parts = key.split('.');
                if (!component[parts[0]]) component[parts[0]] = {};
                component[parts[0]][parts[1]] = value;
            } else {
                component[key] = value;
            }
        });
        if (component.type === 'value-map') {
            component.map = {};
            card.querySelectorAll('.value-map-pairs .flex').forEach(pair => {
                const key = pair.querySelector('.map-key').value;
                const value = parseFloat(pair.querySelector('.map-value').value);
                if (key) component.map[key] = value;
            });
        }
        formula.components.push(component);
    });
    document.querySelectorAll('.modifier-card').forEach(card => {
        const modifier = { operations: [] };
        card.querySelectorAll('input[data-key]').forEach(input => {
            let value = input.value;
            if (input.dataset.key === 'conditionValue') {
                if (value.toLowerCase() === 'true') value = true;
                else if (value.toLowerCase() === 'false') value = false;
                else if (!isNaN(parseFloat(value))) value = parseFloat(value);
            }
            modifier[input.dataset.key] = value;
        });
        card.querySelectorAll('.operations-list .flex').forEach(op => {
            const operator = op.querySelector('.op-operator').value;
            const value = parseFloat(op.querySelector('.op-value').value);
            if (operator && !isNaN(value)) {
                modifier.operations.push({ operator, value });
            }
        });
        formula.postCalculationModifiers.push(modifier);
    });
    if (formula.postCalculationModifiers.length === 0) {
        delete formula.postCalculationModifiers;
    }
    const jsonOutput = document.getElementById('json-output');
    const jsonString = JSON.stringify(formula, null, 2);
    jsonOutput.value = jsonString;
    jsonOutput.select();
    try {
        document.execCommand('copy');
        const copyStatus = document.getElementById('copy-status');
        copyStatus.textContent = "Copied to clipboard!";
        setTimeout(() => { copyStatus.textContent = ""; }, 2000);
    } catch (err) {
        console.error('Failed to copy JSON: ', err);
    }
}
