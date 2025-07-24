/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */

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
 * Reads the 'ldaDaysOut' setting and populates the input field.
 * Defaults to 6 if no setting is found.
 */
function loadSettings() {
  const ldaDaysOut = Office.context.document.settings.get("ldaDaysOut");
  if (ldaDaysOut !== null && ldaDaysOut !== undefined) {
    document.getElementById("ldaDaysOut").value = ldaDaysOut;
  } else {
    // Default to 6 if no setting is saved yet.
    document.getElementById("ldaDaysOut").value = 6;
  }
}

/**
 * Saves the value from the 'ldaDaysOut' input field to the document settings.
 */
function saveSettings() {
  const ldaDaysOutValue = document.getElementById("ldaDaysOut").value;
  const ldaDaysOut = parseInt(ldaDaysOutValue, 10);
  const statusNotification = document.getElementById("status-notification");

  if (!isNaN(ldaDaysOut) && ldaDaysOut >= 0) {
    Office.context.document.settings.set("ldaDaysOut", ldaDaysOut);
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
