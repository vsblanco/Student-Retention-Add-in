/*
 * background-service.js
 *
 * Main entry point and background service for ribbon button commands.
 *
 * This file:
 * - Registers all ribbon button actions with Office.actions
 * - Runs continuously as a background script (even when taskpanes are closed)
 * - Manages Chrome Extension Service communication
 * - Handles Master List data synchronization
 */
import { toggleHighlight, transferData, createSendToCallQueue } from './ribbon-actions.js';
import chromeExtensionService from '../react/src/services/chromeExtensionService.js';
import { CONSTANTS, findColumnIndex, normalizeName, formatToLastFirst, parseDate } from './shared-utilities.js';

// Batch size for chunked read/write operations to avoid Excel's ~5MB payload limit.
// 500 rows is consistent with ldaProcessor.js and safely under the limit for ~37 columns.
const BATCH_SIZE = 500;
// Smaller batch size for formatting operations (cell colors).
const FORMAT_BATCH_SIZE = 50;

/**
 * Sends an import status update back to the Chrome extension via window.postMessage.
 * The extension's content script (excelConnector.js) picks this up and forwards it.
 * @param {string} status - Status key (e.g., 'received', 'writing_data', 'complete', 'error')
 * @param {string} message - Human-readable status message
 * @param {number} [progress] - Optional progress percentage (0-100)
 * @param {string} [error] - Optional error message (for error status)
 */
function sendImportStatus(status, message, progress, error) {
    try {
        window.postMessage({
            type: "SRK_IMPORT_STATUS",
            data: {
                status: status,
                message: message,
                ...(progress != null && { progress }),
                ...(error && { error })
            }
        }, "*");
    } catch (e) {
        console.warn("Failed to send import status:", e);
    }
}

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

// PLACEHOLDER_BACKGROUND_SERVICE