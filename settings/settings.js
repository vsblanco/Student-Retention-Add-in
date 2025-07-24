/* global document, Office */

const SETTINGS_KEY = "studentRetentionSettings";

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    // Determine if the user's version of Office supports all the Office.js APIs that are used in the task pane.
    if (!Office.context.requirements.isSetSupported("ExcelApi", "1.7")) {
      console.log("Sorry. The add-in uses Excel.js APIs that are not available in your version of Office.");
    }

    // Assign event handlers and other initialization logic.
    document.getElementById("save-settings").addEventListener("click", saveSettings);
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
            return JSON.parse(settingsString);
        } catch (e) {
            console.error("Error parsing settings:", e);
            // If parsing fails, return a default structure
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
  }
}

/**
 * Saves the current settings from the form to the document.
 */
function saveSettings() {
  const settings = getSettings();
  const daysOut = document.getElementById("ldaDaysOut").value;

  if (!settings.createlda) {
    settings.createlda = {};
  }
  settings.createlda.daysOutFilter = parseInt(daysOut, 10);

  Office.context.document.settings.set(SETTINGS_KEY, JSON.stringify(settings));
  Office.context.document.settings.saveAsync(function (asyncResult) {
    const status = document.getElementById("status-notification");
    if (asyncResult.status == Office.AsyncResultStatus.Failed) {
      console.error('Settings save failed. Error: ' + asyncResult.error.message);
      status.textContent = "Error saving settings.";
      status.className = "error";
    } else {
      console.log('Settings saved successfully');
      status.textContent = "Settings saved successfully!";
      status.className = "success";
    }
    status.style.display = "block";
    setTimeout(() => {
        status.style.display = "none";
    }, 3000);
  });
}
