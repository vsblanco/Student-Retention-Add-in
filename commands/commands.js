/*
 * This file is the main entry point for all ribbon button commands.
 * It imports functions from other modules and associates them with Office actions.
 *
 * This file also serves as a background script that runs continuously,
 * even when taskpanes are closed, making it ideal for services that need
 * to always be listening (like the Chrome Extension Service).
 */
import { openImportDialog } from './import.js';
import { openCreateLdaDialog } from './lda.js';
import { toggleHighlight, transferData } from './actions.js';
import chromeExtensionService from '../react/src/services/chromeExtensionService.js';

// This function is required for the Analytics button, even if it does nothing,
// because the manifest uses a ShowTaskpane action.
function openAnalyticsPane(event) {
    event.completed();
}

// Called automatically when a new document is opened
// This ensures the background service starts on document load
function onDocumentOpen(event) {
    console.log("Document opened - background services are active");
    event.completed();
}

// Register ribbon button commands
Office.actions.associate("toggleHighlight", toggleHighlight);
Office.actions.associate("openImportDialog", openImportDialog);
Office.actions.associate("transferData", transferData);
Office.actions.associate("openCreateLdaDialog", openCreateLdaDialog);
Office.actions.associate("openAnalyticsPane", openAnalyticsPane);

// Register autoload event handler
Office.actions.associate("onDocumentOpen", onDocumentOpen);

// Initialize background services when Office is ready
Office.onReady(() => {
  console.log("Background script initialized - starting Chrome Extension Service");

  // Start extension detection
  // This will continuously ping to detect the Chrome extension
  chromeExtensionService.startPinging();

  // Start keep-alive heartbeat to prevent extension from going dormant
  chromeExtensionService.startKeepAlive();

  // Add a listener for extension events
  chromeExtensionService.addListener((event) => {
    console.log("Background: Chrome Extension event:", event);

    // Handle different event types
    if (event.type === "installed") {
      console.log("Background: Chrome extension detected and ready!");
    } else if (event.type === "message") {
      console.log("Background: Received message from extension:", event.data);
    }
  });

  console.log("Chrome Extension Service is now running in the background");
});
