Office.onReady((info) => {
    if (info.host === Office.HostType.Excel) {
        document.addEventListener("DOMContentLoaded", run);
    }
});

// --- STATE ---
let currentFormula = null;
let sessionCustomFormulas = {}; // Store multiple custom formulas
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
const RISK_INDEX_COLUMN_NAME = "Risk Index";
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
    "Course Zero Assignments": ["course zero assignments"],
    "Student Name": ["student name", "name", "studentname"]
};
const INFO_CONTENT = {
    modelName: { title: "Model Name", content: "The public-facing name of your formula. This is what will appear in the dropdown list for selection." },
    author: { title: "Author", content: "The person or team who created this model. This helps track ownership and provides contact info for questions." },
    description: { title: "Description", content: "A brief summary of what this model is designed to measure and what makes it unique. This helps users choose the right model for their needs." },
    version: { title: "Version", content: "A version number (e.g., 1.0, 1.1) to help track changes and updates to the model over time." },
    maxScore: { title: "Maximum Score", content: "The highest possible score a student can get with this model. This is typically the sum of all component weights." },
    componentName: { title: "Component Name", content: "The internal, unique identifier for this component (e.g., 'currentGrade'). It should be in camelCase and contain no spaces." },
    componentDisplayName: { title: "Component Display Name", content: "The user-friendly name for this component that will be shown in the formula viewer (e.g., 'Current Grade')." },
    componentWeight: { title: "Weight (Points)", content: "The maximum number of points this component can contribute to the total risk score. The sum of all weights usually equals the 'Maximum Score'." },
    componentCalcType: { title: "Calculation Type", content: `
        Determines how the risk score for this component is calculated.
        <ul class='list-disc list-inside mt-2'>
            <li><strong>value-map:</strong> Assigns specific points based on a text value (e.g., 'A' grade = 0 points).</li>
            <li><strong>linear:</strong> Score increases directly with the value in the source column. (e.g., more 'Days Out' = more points).</li>
            <li><strong>linear-multiplier:</strong> Same as linear, but the value is multiplied by a factor first.</li>
            <li><strong>inverse-linear:</strong> Score decreases as the value in the source column increases (e.g., a higher 'GPA' = fewer points).</li>
            <li><strong>inverse-ratio:</strong> Calculates risk based on the inverse of a ratio between two columns.</li>
        </ul>`
    },
    componentCalcLogic: { title: "Calculation Logic", content: "A plain-text description of how this component works. This text is displayed in the formula viewer to help users understand the model." },
    sourceColumn: { title: "Source Column", content: "The exact name of the column header in your 'Master List' sheet that contains the data for this calculation." },
    valueMappings: { title: "Value Mappings", content: "Define the specific points for each possible value. The 'Key' is the text from the Excel cell (e.g., 'A', 'B', 'C'), and the 'Value' is the number of points to assign." }
};

/**
 * Main function to initialize the task pane.
 */
async function run() {
    // Viewer listeners
    document.getElementById('formula-selector').addEventListener('change', handleFormulaSelection);
    document.getElementById('check-compatibility-button').addEventListener('click', () => checkCompatibility(currentFormula, 'compatibility-results-container'));
    document.getElementById('import-risk-index-btn').addEventListener('click', importRiskIndex);
    
    // Builder listeners
    document.getElementById('add-component-btn').addEventListener('click', addComponent);
    document.getElementById('add-modifier-btn').addEventListener('click', addModifier);
    document.getElementById('generate-json-btn').addEventListener('click', generateAndCopyJson);
    document.getElementById('builder-check-compatibility-btn').addEventListener('click', () => {
        const formula = buildFormulaFromUI();
        checkCompatibility(formula, 'builder-compatibility-results-container');
    });
    document.getElementById('builder-import-btn').addEventListener('click', importFormulaFromBuilder);

    // View toggler
    document.getElementById('toggle-view-btn').addEventListener('click', toggleView);
    document.getElementById('edit-custom-formula-btn').addEventListener('click', openEditModal);

    // Modal Listeners
    document.getElementById('load-custom-formula-btn').addEventListener('click', handleCustomFormulaLoad);
    document.getElementById('cancel-custom-formula-btn').addEventListener('click', () => {
        document.getElementById('custom-formula-modal').classList.add('hidden');
        document.getElementById('formula-selector').value = ''; 
    });
    document.getElementById('cancel-edit-formula-btn').addEventListener('click', () => {
        document.getElementById('edit-formula-modal').classList.add('hidden');
    });

    // Info Popup Listeners
    document.body.addEventListener('click', (event) => {
        if (event.target.classList.contains('info-btn')) {
            const infoKey = event.target.dataset.info;
            const infoData = INFO_CONTENT[infoKey];
            if (infoData) {
                showInfoPopup(infoData.title, infoData.content);
            }
        }
    });
    document.getElementById('info-popup-close-btn').addEventListener('click', () => {
        document.getElementById('info-popup-modal').classList.add('hidden');
    });

    // Event delegation for expandable example scores
    document.body.addEventListener('click', (event) => {
        const scoreHeader = event.target.closest('.score-example-header');
        if (scoreHeader) {
            const details = scoreHeader.nextElementSibling;
            details.classList.toggle('hidden');
        }
    });

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
        viewerMode.classList.remove('hidden');
        builderMode.classList.add('hidden');
        toggleBtn.textContent = 'Formula Builder';
        mainTitle.textContent = 'Risk Index Models';
    } else {
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
    const divider = document.createElement('option');
    divider.disabled = true;
    divider.textContent = '──────────';
    selector.appendChild(divider);
    const customOption = document.createElement('option');
    customOption.value = 'custom';
    customOption.textContent = 'Upload Custom Formula...';
    selector.appendChild(customOption);
}

function handleFormulaSelection() {
    const selector = document.getElementById('formula-selector');
    const selectedValue = selector.value;
    if (selectedValue === 'custom') {
        document.getElementById('custom-json-input').value = '';
        document.getElementById('custom-formula-status').textContent = '';
        document.getElementById('custom-formula-modal').classList.remove('hidden');
    } else if (selectedValue.startsWith('session-custom-')) {
        const modelName = selectedValue.replace('session-custom-', '');
        displayFormula(null, sessionCustomFormulas[modelName]);
    } else {
        displayFormula(selectedValue);
    }
}

function handleCustomFormulaLoad() {
    const jsonInput = document.getElementById('custom-json-input');
    const statusEl = document.getElementById('custom-formula-status');
    try {
        const formulaObject = JSON.parse(jsonInput.value);
        if (!formulaObject.modelName || !formulaObject.components) {
            throw new Error('JSON is missing required fields like "modelName" or "components".');
        }
        sessionCustomFormulas[formulaObject.modelName] = formulaObject;
        updateCustomFormulaDropdown();
        document.getElementById('formula-selector').value = `session-custom-${formulaObject.modelName}`;
        document.getElementById('custom-formula-modal').classList.add('hidden');
        displayFormula(null, formulaObject);
    } catch (error) {
        statusEl.textContent = `Error: ${error.message}`;
        console.error("Invalid JSON:", error);
    }
}

function updateCustomFormulaDropdown() {
    const selector = document.getElementById('formula-selector');
    // Remove existing custom options
    selector.querySelectorAll('option[value^="session-custom-"]').forEach(opt => opt.remove());

    Object.values(sessionCustomFormulas).forEach(formula => {
        const customOption = document.createElement('option');
        customOption.value = `session-custom-${formula.modelName}`;
        customOption.textContent = `(Custom) ${formula.modelName}`;
        selector.insertBefore(customOption, selector.querySelector('option[value="custom"]'));
    });

    // Show/hide edit button
    const editBtn = document.getElementById('edit-custom-formula-btn');
    if(Object.keys(sessionCustomFormulas).length > 0) {
        editBtn.classList.remove('hidden');
    } else {
        editBtn.classList.add('hidden');
    }
}

async function displayFormula(fileName, formulaObject = null) {
    const container = document.getElementById('formula-details-container');
    const spinner = document.getElementById('loading-spinner');
    const compatibilitySection = document.getElementById('compatibility-section');
    const resultsContainer = document.getElementById('compatibility-results-container');
    currentFormula = null;
    resultsContainer.innerHTML = '';
    document.getElementById('import-risk-index-btn').classList.add('hidden'); // Hide import button initially
    if (!fileName && !formulaObject) {
        container.innerHTML = '<p class="text-center text-gray-500">Please select a formula to view its details.</p>';
        compatibilitySection.classList.add('hidden');
        return;
    }
    container.classList.add('hidden');
    spinner.classList.remove('hidden');
    compatibilitySection.classList.add('hidden');
    container.innerHTML = '';
    try {
        currentFormula = formulaObject ? formulaObject : await (await fetch(`${FORMULAS_PATH}${fileName}`)).json();
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

// ########################################
// ### COMPATIBILITY & EXAMPLE CHECKER  ###
// ########################################

async function checkCompatibility(formulaToCheck, resultsContainerId) {
    if (!formulaToCheck) return;
    const resultsContainer = document.getElementById(resultsContainerId);
    resultsContainer.innerHTML = '<p class="text-center text-gray-500">Checking...</p>';
    document.getElementById('import-risk-index-btn').classList.add('hidden');
    try {
        const requiredColumns = new Set();
        (formulaToCheck.components || []).forEach(c => {
            if (c.column) requiredColumns.add(c.column);
            if (c.columns) {
                if (c.columns.numerator) requiredColumns.add(c.columns.numerator);
                if (c.columns.denominator) requiredColumns.add(c.columns.denominator);
            }
        });
        (formulaToCheck.postCalculationModifiers || []).forEach(m => {
            if (m.conditionColumn) requiredColumns.add(m.conditionColumn);
        });
        requiredColumns.add("Student Name"); // Always need student name for examples

        const { data, headerMap } = await Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getItem(MASTER_LIST_SHEET);
            const usedRange = sheet.getUsedRange(true);
            usedRange.load("values");
            await context.sync();
            
            const sheetData = usedRange.values;
            const headers = sheetData[0].map(h => String(h || '').toLowerCase());
            
            const map = {};
            requiredColumns.forEach(reqCol => {
                const aliases = COLUMN_ALIASES[reqCol] || [reqCol.toLowerCase()];
                const foundIndex = headers.findIndex(header => aliases.includes(header));
                if(foundIndex !== -1) {
                    map[reqCol] = foundIndex;
                }
            });
            return { data: sheetData, headerMap: map };
        });

        const studentDataRows = data.slice(1);
        const columnCheckResults = [];
        let allFound = true;
        requiredColumns.forEach(reqCol => {
            const found = headerMap[reqCol] !== undefined;
            if (!found) allFound = false;
            columnCheckResults.push({ name: reqCol, found });
        });
        
        let exampleScores = [];
        if (allFound && studentDataRows.length > 0) {
            document.getElementById('import-risk-index-btn').classList.remove('hidden'); // Show import button if compatible
            const sampleSize = Math.min(5, studentDataRows.length);
            const randomIndices = [...Array(studentDataRows.length).keys()].sort(() => 0.5 - Math.random()).slice(0, sampleSize);
            
            randomIndices.forEach(index => {
                const studentRow = studentDataRows[index];
                const { score, breakdown } = calculateRiskScore(studentRow, formulaToCheck, headerMap);
                const studentName = studentRow[headerMap["Student Name"]] || `Row ${index + 2}`;
                exampleScores.push({ name: studentName, score, breakdown });
            });
        }

        renderCompatibilityResults(columnCheckResults, allFound, exampleScores, resultsContainerId, formulaToCheck.maxScore);

    } catch (error) {
        let errorMessage = `Error during compatibility check: ${error.message}`;
        if (error.code === 'ItemNotFound') {
            errorMessage = `Error: The "${MASTER_LIST_SHEET}" sheet could not be found.`;
        }
        resultsContainer.innerHTML = `<p class="text-center text-red-500">${errorMessage}</p>`;
        console.error(error);
    }
}

function calculateRiskScore(studentRow, formula, headerMap) {
    let totalScore = 0;
    const breakdown = [];

    // Calculate component scores
    (formula.components || []).forEach(comp => {
        let componentScore = 0;
        const weight = comp.weight || 0;
        
        try {
            switch(comp.type) {
                case 'value-map': {
                    const value = studentRow[headerMap[comp.column]];
                    componentScore = comp.map[value] !== undefined ? comp.map[value] : 0;
                    break;
                }
                case 'linear': {
                    const value = parseFloat(studentRow[headerMap[comp.column]]) || 0;
                    componentScore = Math.min(value, weight);
                    break;
                }
                case 'linear-multiplier': {
                    const value = parseFloat(studentRow[headerMap[comp.column]]) || 0;
                    const multiplier = comp.multiplier || 1;
                    componentScore = Math.min(value * multiplier, weight);
                    break;
                }
                case 'inverse-linear': {
                     const value = parseFloat(studentRow[headerMap[comp.column]]) || 0;
                     // Assuming value is a percentage for this calculation
                     componentScore = (1 - (value / 100)) * weight;
                     break;
                }
                case 'inverse-ratio': {
                    const num = parseFloat(studentRow[headerMap[comp.columns.numerator]]) || 0;
                    const den = parseFloat(studentRow[headerMap[comp.columns.denominator]]) || 0;
                    if (num > 0) {
                        componentScore = (1 - (den / num)) * weight;
                    }
                    break;
                }
            }
        } catch (e) { console.error(`Error calculating component ${comp.name}:`, e); }

        const finalComponentScore = Math.round(Math.max(0, componentScore));
        breakdown.push({ name: comp.displayName || comp.name, score: finalComponentScore });
        totalScore += finalComponentScore;
    });

    // Apply modifiers
    (formula.postCalculationModifiers || []).forEach(mod => {
        const conditionValue = studentRow[headerMap[mod.conditionColumn]];
        let conditionMet = String(conditionValue).toLowerCase() === String(mod.conditionValue).toLowerCase();
        if (typeof mod.conditionValue === 'boolean') {
            conditionMet = (String(conditionValue).toLowerCase() === 'true') === mod.conditionValue;
        }

        if (conditionMet) {
            (mod.operations || []).forEach(op => {
                switch(op.operator) {
                    case '+': totalScore += op.value; break;
                    case '-': totalScore -= op.value; break;
                    case '*': totalScore *= op.value; break;
                    case '/': if(op.value !== 0) totalScore /= op.value; break;
                }
            });
        }
    });

    return { score: Math.round(Math.min(totalScore, formula.maxScore || 100)), breakdown };
}

function getScoreColor(score, maxScore) {
    if (maxScore === 0) return 'hsl(120, 70%, 45%)'; // Default to green if maxScore is 0
    const percentage = Math.min(score / maxScore, 1); // Cap at 100%
    // Hue goes from 120 (green) down to 0 (red)
    const hue = 120 * (1 - percentage);
    // Use HSL for easy color interpolation
    return `hsl(${hue}, 70%, 45%)`;
}

function renderCompatibilityResults(results, allFound, exampleScores, containerId, maxScore) {
    const container = document.getElementById(containerId);
    let overallStatus = allFound
        ? `<div class="p-3 mb-3 bg-green-100 text-green-800 rounded-lg text-sm font-semibold">✔ All required columns found. This model is compatible.</div>`
        : `<div class="p-3 mb-3 bg-red-100 text-red-800 rounded-lg text-sm font-semibold">❌ Missing required columns. This model may not run correctly.</div>`;
    
    let listHtml = '<div><h4 class="text-md font-semibold text-gray-700 mb-2">Column Requirements</h4><ul class="space-y-1 text-sm">';
    results.forEach(res => {
        const icon = res.found ? '<span class="text-green-500 font-bold mr-2">✔</span>' : '<span class="text-red-500 font-bold mr-2">❌</span>';
        const textColor = res.found ? 'text-gray-700' : 'text-red-700 font-semibold';
        listHtml += `<li class="p-2 bg-gray-50 rounded-md flex items-center ${textColor}">${icon} ${res.name}</li>`;
    });
    listHtml += '</ul></div>';

    let examplesHtml = '';
    if (exampleScores && exampleScores.length > 0) {
        examplesHtml = '<div class="mt-4"><h4 class="text-md font-semibold text-gray-700 mb-2">Example Scores (Click to expand)</h4><ul class="space-y-1 text-sm">';
        exampleScores.forEach(ex => {
            const scoreColor = getScoreColor(ex.score, maxScore || 100);
            let breakdownHtml = '<ul class="mt-2 border-t pt-2 space-y-1">';
            ex.breakdown.forEach(item => {
                breakdownHtml += `<li class="flex justify-between text-xs text-gray-600"><span>${item.name}</span><span class="font-medium">${item.score} pts</span></li>`;
            });
            breakdownHtml += '</ul>';

            examplesHtml += `
                <li class="p-2 bg-gray-50 rounded-md">
                    <div class="score-example-header flex justify-between items-center cursor-pointer">
                        <span>${ex.name}</span>
                        <span class="font-bold text-white px-2 py-1 rounded" style="background-color: ${scoreColor};">${ex.score} pts</span>
                    </div>
                    <div class="score-details hidden mt-2">
                        ${breakdownHtml}
                    </div>
                </li>`;
        });
        examplesHtml += '</ul></div>';
    }

    container.innerHTML = overallStatus + listHtml + examplesHtml;
}


// ####################################
// ### FORMULA BUILDER FUNCTIONS    ###
// ####################################

function openEditModal() {
    const listContainer = document.getElementById('edit-formula-list');
    listContainer.innerHTML = ''; // Clear previous list
    const customFormulaNames = Object.keys(sessionCustomFormulas);

    if (customFormulaNames.length === 0) {
        alert("No custom formulas in this session to edit.");
        return;
    }

    customFormulaNames.forEach(name => {
        const button = document.createElement('button');
        button.className = 'w-full text-left p-2 bg-gray-100 rounded-md hover:bg-blue-100';
        button.textContent = name;
        button.onclick = () => {
            loadFormulaIntoBuilder(name);
            document.getElementById('edit-formula-modal').classList.add('hidden');
        };
        listContainer.appendChild(button);
    });

    document.getElementById('edit-formula-modal').classList.remove('hidden');
}

function loadFormulaIntoBuilder(modelName) {
    const formula = sessionCustomFormulas[modelName];
    if (!formula) return;

    // Switch to builder view if not already there
    if (document.getElementById('viewer-mode').style.display !== 'none') {
        toggleView();
    }
    
    // Populate general info
    document.getElementById('modelName').value = formula.modelName || '';
    document.getElementById('author').value = formula.author || '';
    document.getElementById('description').value = formula.description || '';
    document.getElementById('version').value = formula.version || '1.0';
    document.getElementById('maxScore').value = formula.maxScore || 100;

    // Populate components
    const componentsContainer = document.getElementById('components-container');
    componentsContainer.innerHTML = '';
    componentIdCounter = 0;
    (formula.components || []).forEach(compData => {
        addComponent(); // Adds a blank card
        const newCard = document.getElementById(`comp-${componentIdCounter - 1}`);
        // Now populate the new card
        newCard.querySelector('[data-key="name"]').value = compData.name || '';
        newCard.querySelector('[data-key="displayName"]').value = compData.displayName || '';
        newCard.querySelector('[data-key="weight"]').value = compData.weight || 10;
        const typeSelect = newCard.querySelector('[data-key="type"]');
        typeSelect.value = compData.type || 'linear';
        handleTypeChange({target: typeSelect}); // Trigger UI update for the selected type
        
        // Populate type-specific fields
        if(compData.column) newCard.querySelector('[data-key="column"]').value = compData.column;
        if(compData.multiplier) newCard.querySelector('[data-key="multiplier"]').value = compData.multiplier;
        if(compData.columns) {
            if(compData.columns.numerator) newCard.querySelector('[data-key="columns.numerator"]').value = compData.columns.numerator;
            if(compData.columns.denominator) newCard.querySelector('[data-key="columns.denominator"]').value = compData.columns.denominator;
        }
        if(compData.type === 'value-map' && compData.map) {
            const pairsContainer = newCard.querySelector('.value-map-pairs');
            pairsContainer.innerHTML = ''; // Clear the default pair
            Object.entries(compData.map).forEach(([key, value]) => {
                addValueMapPair(pairsContainer, key, value);
            });
        }
    });

    // Populate modifiers
    const modifiersContainer = document.getElementById('modifiers-container');
    modifiersContainer.innerHTML = '';
    modifierIdCounter = 0;
    (formula.postCalculationModifiers || []).forEach(modData => {
        addModifier();
        const newCard = document.getElementById(`mod-${modifierIdCounter - 1}`);
        // Populate fields
        newCard.querySelector('[data-key="name"]').value = modData.name || '';
        newCard.querySelector('[data-key="displayName"]').value = modData.displayName || '';
        newCard.querySelector('[data-key="conditionColumn"]').value = modData.conditionColumn || '';
        newCard.querySelector('[data-key="conditionValue"]').value = modData.conditionValue || '';
        newCard.querySelector('[data-key="calculation"]').value = modData.calculation || '';
        
        const opsContainer = newCard.querySelector('.operations-list');
        opsContainer.innerHTML = ''; // Clear default
        (modData.operations || []).forEach(op => {
            addOperation(opsContainer, op.operator, op.value);
        });
    });
}


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
            <div><label class="block font-medium">Name <span class="info-btn" data-info="componentName">(i)</span></label><input type="text" data-key="name" class="w-full p-1 border rounded" placeholder="e.g., currentGrade"></div>
            <div><label class="block font-medium">Display Name <span class="info-btn" data-info="componentDisplayName">(i)</span></label><input type="text" data-key="displayName" class="w-full p-1 border rounded" placeholder="e.g., Current Grade"></div>
            <div><label class="block font-medium">Weight (Points) <span class="info-btn" data-info="componentWeight">(i)</span></label><input type="number" data-key="weight" class="w-full p-1 border rounded" value="10"></div>
            <div><label class="block font-medium">Calculation Type <span class="info-btn" data-info="componentCalcType">(i)</span></label><select data-key="type" class="w-full p-1 border rounded component-type-select">${COMPONENT_TYPES.map(t => `<option value="${t}">${t}</option>`).join('')}</select></div>
            <div class="md:col-span-2"><label class="block font-medium">Calculation Logic <span class="info-btn" data-info="componentCalcLogic">(i)</span></label><input type="text" data-key="calculation" class="w-full p-1 border rounded" placeholder="e.g., Points assigned based on letter grade."></div>
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
        configHtml += `<div><label class="block font-medium">Source Column <span class="info-btn" data-info="sourceColumn">(i)</span></label><input type="text" data-key="column" class="w-full p-1 border rounded" placeholder="e.g., Grade"></div>`;
    }
    switch (selectedType) {
        case 'linear-multiplier':
            configHtml += `<div><label class="block font-medium">Multiplier</label><input type="number" data-key="multiplier" class="w-full p-1 border rounded" value="1"></div>`;
            break;
        case 'inverse-ratio':
            configHtml += `<div><label class="block font-medium">Numerator Column</label><input type="text" data-key="columns.numerator" class="w-full p-1 border rounded" placeholder="e.g., Contact Attempts"></div><div><label class="block font-medium">Denominator Column</label><input type="text" data-key="columns.denominator" class="w-full p-1 border rounded" placeholder="e.g., Times Contacted"></div>`;
            break;
        case 'value-map':
            configHtml += `<div class="md:col-span-2"><label class="block font-medium">Value Mappings <span class="info-btn" data-info="valueMappings">(i)</span></label><div class="value-map-pairs space-y-2 mt-1"></div><button class="add-map-pair-btn mt-2 text-xs px-2 py-1 bg-gray-200 rounded hover:bg-gray-300">+ Add Pair</button></div>`;
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

function addValueMapPair(container, key = '', value = '') {
    const pairDiv = document.createElement('div');
    pairDiv.className = 'flex items-center gap-2';
    pairDiv.innerHTML = `<input type="text" class="w-full p-1 border rounded map-key" placeholder="Key (e.g., A)" value="${key}"><span class="font-bold">:</span><input type="number" class="w-full p-1 border rounded map-value" placeholder="Value (e.g., 0)" value="${value}"><button class="remove-pair-btn text-red-500 hover:text-red-700">&times;</button>`;
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

function addOperation(container, operator = '+', value = '') {
    const opDiv = document.createElement('div');
    opDiv.className = 'flex items-center gap-2';
    opDiv.innerHTML = `<select class="w-1/3 p-1 border rounded op-operator">${MODIFIER_OPERATORS.map(op => `<option value="${op}" ${op === operator ? 'selected' : ''}>${op}</option>`).join('')}</select><input type="number" class="w-2/3 p-1 border rounded op-value" placeholder="Value" value="${value}"><button class="remove-op-btn text-red-500 hover:text-red-700">&times;</button>`;
    opDiv.querySelector('.remove-op-btn').addEventListener('click', () => opDiv.remove());
    container.appendChild(opDiv);
}

function buildFormulaFromUI() {
    const formula = {
        modelName: document.getElementById('modelName').value || "Untitled Model",
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
    return formula;
}

function generateAndCopyJson() {
    const formula = buildFormulaFromUI();
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

function importFormulaFromBuilder() {
    const formula = buildFormulaFromUI();
    if (!formula.modelName) {
        alert("Please provide a Model Name before importing.");
        return;
    }
    // Remove old version if name is the same but case might be different
    const oldKey = Object.keys(sessionCustomFormulas).find(key => key.toLowerCase() === formula.modelName.toLowerCase());
    if(oldKey) {
        delete sessionCustomFormulas[oldKey];
    }

    sessionCustomFormulas[formula.modelName] = formula;
    updateCustomFormulaDropdown();
    toggleView(); // Switch back to viewer
    document.getElementById('formula-selector').value = `session-custom-${formula.modelName}`;
    displayFormula(null, formula); // Display the new formula
}

// ####################################
// ### INFO POPUP FUNCTIONS         ###
// ####################################

function showInfoPopup(title, content) {
    document.getElementById('info-popup-title').textContent = title;
    document.getElementById('info-popup-content').innerHTML = content;
    document.getElementById('info-popup-modal').classList.remove('hidden');
}

// ####################################
// ### RISK INDEX IMPORT FUNCTIONS  ###
// ####################################

async function importRiskIndex() {
    if (!currentFormula) {
        alert("Please select a formula first.");
        return;
    }

    try {
        await Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getItem(MASTER_LIST_SHEET);
            const usedRange = sheet.getUsedRange(true);
            usedRange.load(["values", "columnCount", "rowCount"]);
            await context.sync();

            const headers = usedRange.values[0].map(h => String(h || '').toLowerCase());
            let riskIndexCol = headers.indexOf(RISK_INDEX_COLUMN_NAME.toLowerCase());

            if (riskIndexCol === -1) {
                const userConfirmed = await showConfirmationModal(`The "${RISK_INDEX_COLUMN_NAME}" column was not found. Would you like to create it?`);
                if (userConfirmed) {
                    riskIndexCol = usedRange.columnCount;
                    sheet.getRangeByIndexes(0, riskIndexCol, 1, 1).values = [[RISK_INDEX_COLUMN_NAME]];
                } else {
                    return; // User cancelled
                }
            }
            
            // Re-fetch data and build header map now that column is guaranteed
            const { data, headerMap } = await getSheetDataWithHeaderMap();
            const scores = data.slice(1).map(studentRow => {
                const { score } = calculateRiskScore(studentRow, currentFormula, headerMap);
                return [score];
            });

            if (scores.length > 0) {
                const targetRange = sheet.getRangeByIndexes(1, riskIndexCol, scores.length, 1);
                targetRange.values = scores;
                
                // Add conditional formatting
                const formatRange = sheet.getRangeByIndexes(1, riskIndexCol, usedRange.rowCount - 1, 1);
                formatRange.conditionalFormats.clearAll();
                const conditionalFormat = formatRange.conditionalFormats.add(Excel.ConditionalFormatType.colorScale);
                conditionalFormat.colorScale.criteria = {
                    minimum: { type: Excel.ConditionalFormatRuleType.lowestValue, color: "#63BE7B" }, // Green
                    midpoint: { type: Excel.ConditionalFormatRuleType.percentile, percentile: 50, color: "#FFEB84" }, // Yellow
                    maximum: { type: Excel.ConditionalFormatRuleType.highestValue, color: "#F8696B" }  // Red
                };
            }
            
            await context.sync();
        });
        alert("Risk Index scores have been successfully imported and formatted.");
    } catch (error) {
        console.error("Error importing risk index:", error);
        alert(`An error occurred: ${error.message}`);
    }
}

async function getSheetDataWithHeaderMap() {
    return await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getItem(MASTER_LIST_SHEET);
        const usedRange = sheet.getUsedRange(true);
        usedRange.load("values");
        await context.sync();
        
        const sheetData = usedRange.values;
        const headers = sheetData[0].map(h => String(h || '').toLowerCase());
        
        const map = {};
        Object.keys(COLUMN_ALIASES).forEach(key => {
            const aliases = COLUMN_ALIASES[key];
            const foundIndex = headers.findIndex(header => aliases.includes(header));
            if(foundIndex !== -1) {
                map[key] = foundIndex;
            }
        });
        return { data: sheetData, headerMap: map };
    });
}


function showConfirmationModal(message) {
    return new Promise((resolve) => {
        document.getElementById('confirmation-message').textContent = message;
        const modal = document.getElementById('confirmation-modal');
        modal.classList.remove('hidden');

        const okBtn = document.getElementById('confirmation-ok-btn');
        const cancelBtn = document.getElementById('confirmation-cancel-btn');

        const cleanup = () => {
            modal.classList.add('hidden');
            okBtn.replaceWith(okBtn.cloneNode(true));
            cancelBtn.replaceWith(cancelBtn.cloneNode(true));
        };

        okBtn.addEventListener('click', () => {
            cleanup();
            resolve(true);
        });

        cancelBtn.addEventListener('click', () => {
            cleanup();
            resolve(false);
        });
    });
}
