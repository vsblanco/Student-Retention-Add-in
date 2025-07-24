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
            return { transferData: {} };
        }
    }
    return { transferData: {} };
}
