/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

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


/**
 * Reads settings and populates the input fields.
 */
function loadSettings() {
  const settings = getSettings();
  const ldaDaysOut = settings.transferData && settings.transferData.daysOutFilter !== undefined 
    ? settings.transferData.daysOutFilter 
    : 6; // Default to 6 if not set
  document.getElementById("ldaDaysOut").value = ldaDaysOut;
}

/**
 * Saves settings from the input fields to the document settings.
 */
function saveSettings() {
  const ldaDaysOutValue = document.getElementById("ldaDaysOut").value;
  const ldaDaysOut = parseInt(ldaDaysOutValue, 10);
  const statusNotification = document.getElementById("status-notification");

  if (!isNaN(ldaDaysOut) && ldaDaysOut >= 0) {
    const currentSettings = getSettings();
    if (!currentSettings.transferData) {
        currentSettings.transferData = {};
    }
    currentSettings.transferData.daysOutFilter = ldaDaysOut;
    
    Office.context.document.settings.set(SETTINGS_KEY, JSON.stringify(currentSettings));
    Office.context.document.settings.saveAsync(function (asyncResult) {
      if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
        statusNotification.textContent = "Settings saved successfully!";
        statusNotification.className = "success";
        statusNotification.style.display = "block";
      } else {
        statusNotification.textContent = "Error: " + asyncResult.error.message;
        statusNotification.className = "error";
        statusNotification.style.display = "block";
      }
      // Hide the notification after 3 seconds
      setTimeout(() => {
        statusNotification.style.display = "none";
      }, 3000);
    });
  } else {
    statusNotification.textContent = "Please enter a valid non-negative number for days out.";
    statusNotification.className = "error";
    statusNotification.style.display = "block";
  }
}
