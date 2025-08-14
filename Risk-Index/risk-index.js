Office.onReady((info) => {
    if (info.host === Office.HostType.Excel) {
        document.addEventListener("DOMContentLoaded", run);
    }
});

// --- STATE ---
let currentFormula = null;
let sessionCustomFormulas = {}; 
let componentIdCounter = 0;
let modifierIdCounter = 0;
let currentBuilderStep = 0;

// --- CONSTANTS ---
const FORMULA_FILES = [
    "standard-risk-model.json",
    "re-entry-risk-model.json",
    "engagement-momentum-model.json",
    "Simple-Risk-Index.json"
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
    // ... content from previous version ...
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
    document.getElementById('builder-check-compatibility-btn').addEventListener('click', () => {
        const formula = buildFormulaFromUI();
        checkCompatibility(formula, 'builder-compatibility-results-container');
    });
    document.getElementById('builder-import-btn').addEventListener('click', importFormulaFromBuilder);
    document.getElementById('builder-prev-btn').addEventListener('click', () => navigateBuilder(-1));
    document.getElementById('builder-next-btn').addEventListener('click', () => navigateBuilder(1));
    document.getElementById('builder-json-popup-btn').addEventListener('click', showJsonPopup);
    document.getElementById('builder-tabs').addEventListener('click', (e) => {
        if (e.target.matches('.tab-btn')) {
            const step = parseInt(e.target.dataset.step, 10);
            showBuilderStep(step);
        }
    });

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
    document.getElementById('close-json-modal-btn').addEventListener('click', () => {
        document.getElementById('json-output-modal').classList.add('hidden');
    });
    document.getElementById('copy-json-btn').addEventListener('click', copyJsonToClipboard);

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
    showBuilderStep(0); // Initialize builder to the first step
}

// ... other functions from previous version ...

// ####################################
// ### FORMULA BUILDER FUNCTIONS    ###
// ####################################

function navigateBuilder(direction) {
    const newStep = currentBuilderStep + direction;
    if (newStep >= 0 && newStep < 4) {
        showBuilderStep(newStep);
    }
}

function showBuilderStep(stepIndex) {
    currentBuilderStep = stepIndex;
    
    // Update tabs
    document.querySelectorAll('#builder-tabs .tab-btn').forEach((btn, index) => {
        btn.classList.toggle('active', index === stepIndex);
    });

    // Update sections
    document.querySelectorAll('.builder-section').forEach((section, index) => {
        section.classList.toggle('hidden', index !== stepIndex);
    });

    // Update nav buttons
    document.getElementById('builder-prev-btn').style.visibility = stepIndex === 0 ? 'hidden' : 'visible';
    document.getElementById('builder-next-btn').style.visibility = stepIndex === 3 ? 'hidden' : 'visible';
}

function showJsonPopup() {
    const formula = buildFormulaFromUI();
    const jsonOutput = document.getElementById('json-output');
    jsonOutput.value = JSON.stringify(formula, null, 2);
    document.getElementById('copy-status').textContent = '';
    document.getElementById('json-output-modal').classList.remove('hidden');
}

function copyJsonToClipboard() {
    const jsonOutput = document.getElementById('json-output');
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

// ... rest of the builder functions (addComponent, buildFormulaFromUI, etc.) ...
