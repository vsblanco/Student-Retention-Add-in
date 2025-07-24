/* global document, Office, Sortable */

const SETTINGS_KEY = "studentRetentionSettings";
const DEFAULT_LDA_COLUMNS = ['Assigned', 'StudentName', 'StudentNumber', 'LDA', 'Days Out', 'grade', 'Phone', 'Outreach'];

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    if (!Office.context.requirements.isSetSupported("ExcelApi", "1.7")) {
      console.log("Sorry. The add-in uses Excel.js APIs that are not available in your version of Office.");
    }

    // Assign event handlers and other initialization logic.
    document.getElementById("save-settings").addEventListener("click", saveSettings);
    document.getElementById("lda-column-input").addEventListener("keydown", handleColumnInput);
    document.getElementById("lda-columns-list").addEventListener("click", handleColumnRemove);
    
    loadSettings();
    console.log("Settings task pane initialized.");
  }
});

/**
 * Gets the settings object from document settings.
 * @returns {object} The parsed settings object.
 */
function getSettings() {
    const settingsString = Office.context.document.settings.get(SETTINGS_KEY);
    if (settingsString) {
        try {
            const settings = JSON.parse(settingsString);
            // Ensure the createlda property exists
            if (!settings.createlda) {
                settings.createlda = {};
            }
            return settings;
        } catch (e) {
            console.error("Error parsing settings:", e);
            return { createlda: {} };
        }
    }
    return { createlda: {} };
}

/**
 * Loads settings and populates the form.
 */
function loadSettings() {
  const settings = getSettings();
  if (settings && settings.createlda) {
    document.getElementById("ldaDaysOut").value = settings.createlda.daysOutFilter || 6;
    document.getElementById("includeFailingList").checked = settings.createlda.includeFailingList ?? true;
    
    const ldaColumns = settings.createlda.ldaColumns || DEFAULT_LDA_COLUMNS;
    renderLdaColumns(ldaColumns);
  } else {
    // Default values for a fresh setup
    document.getElementById("includeFailingList").checked = true;
    renderLdaColumns(DEFAULT_LDA_COLUMNS);
  }
}

/**
 * Renders the list of LDA columns as draggable tags.
 * @param {string[]} columns - An array of column names.
 */
function renderLdaColumns(columns) {
    const listElement = document.getElementById("lda-columns-list");
    listElement.innerHTML = ''; // Clear existing columns
    columns.forEach(columnName => {
        const columnTag = createColumnTag(columnName);
        listElement.appendChild(columnTag);
    });

    // Initialize SortableJS for drag-and-drop
    new Sortable(listElement, {
        animation: 150,
        ghostClass: 'bg-blue-100'
    });
}

/**
 * Creates a single draggable column tag element.
 * @param {string} name - The name of the column.
 * @returns {HTMLElement} The created div element for the tag.
 */
function createColumnTag(name) {
    const div = document.createElement('div');
    div.className = "flex items-center bg-gray-200 text-gray-700 text-sm font-medium px-3 py-1 rounded-full cursor-move";
    div.innerHTML = `
        <span class="column-name">${name}</span>
        <button type="button" class="ml-2 text-gray-500 hover:text-gray-700 remove-column-btn" title="Remove column">&times;</button>
    `;
    return div;
}

/**
 * Handles the keydown event on the column input field.
 * Adds a new column when Enter is pressed.
 * @param {KeyboardEvent} event
 */
function handleColumnInput(event) {
    if (event.key === 'Enter') {
        event.preventDefault();
        const input = event.target;
        const columnName = input.value.trim();
        if (columnName) {
            const listElement = document.getElementById("lda-columns-list");
            const newTag = createColumnTag(columnName);
            listElement.appendChild(newTag);
            input.value = '';
        }
    }
}

/**
 * Handles click events on the column list to remove a column tag.
 * @param {MouseEvent} event
 */
function handleColumnRemove(event) {
    if (event.target.classList.contains('remove-column-btn')) {
        event.target.parentElement.remove();
    }
}


/**
 * Saves the current settings from the form to the document.
 */
function saveSettings() {
  const settings = getSettings();
  const daysOut = document.getElementById("ldaDaysOut").value;
  const includeFailing = document.getElementById("includeFailingList").checked;
  
  // Get columns from the DOM
  const ldaColumnsList = document.getElementById("lda-columns-list");
  const columnElements = ldaColumnsList.querySelectorAll('.column-name');
  const ldaColumns = Array.from(columnElements).map(el => el.textContent);

  if (!settings.createlda) {
    settings.createlda = {};
  }
  settings.createlda.daysOutFilter = parseInt(daysOut, 10);
  settings.createlda.includeFailingList = includeFailing;
  settings.createlda.ldaColumns = ldaColumns;

  Office.context.document.settings.set(SETTINGS_KEY, JSON.stringify(settings));
  Office.context.document.settings.saveAsync(function (asyncResult) {
    const status = document.getElementById("status-notification");
    if (asyncResult.status == Office.AsyncResultStatus.Failed) {
      console.error('Settings save failed. Error: ' + asyncResult.error.message);
      status.textContent = "Error saving settings.";
      status.classList.remove("success");
      status.classList.add("error");
    } else {
      console.log('Settings saved successfully');
      status.textContent = "Settings saved successfully!";
      status.classList.remove("error");
      status.classList.add("success");
    }
    setTimeout(() => {
        status.textContent = "";
        status.classList.remove("success", "error");
    }, 3000);
  });
}
