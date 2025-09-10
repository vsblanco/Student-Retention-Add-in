/**
 * state.js
 * * A simple, centralized state management module for the Personalized Email task pane.
 * This module helps to avoid prop-drilling and provides a single source of truth
 * for application-wide data.
 * * The state is intentionally not persistent (it resets on page reload). For persistent
 * data like the Power Automate URL or templates, the data.js module handles storage
 * in Office settings.
 */

/**
 * The initial shape of the application's state.
 * All properties that will be used throughout the app's lifecycle should be
 * defined here with a default value. This prevents accidental creation of
 * state properties and makes the state's structure predictable.
 * * - powerAutomateUrl: The URL for the Power Automate flow.
 * - students: An array of student data objects fetched from the current sheet.
 * - customParameters: An array of custom parameter configurations.
 * - quill: The instance of the Quill editor.
 * - lastFocusedElement: A reference to the last focused input or editor.
 */
const appState = {
    powerAutomateUrl: null,
    students: [],
    customParameters: [],
    quill: null,
    lastFocusedElement: null // FIX: Added the lastFocusedElement property.
};

/**
 * Retrieves a value from the state.
 * For simplicity, this returns the entire state object.
 * * @returns {object} The current state object.
 */
export function getState() {
    return { ...appState };
}

/**
 * Updates a value in the state.
 * This function enforces that only pre-defined keys in the initial appState
 * can be updated, preventing accidental state pollution.
 * * @param {string} key The key of the state property to update.
 * @param {*} value The new value for the property.
 */
export function updateState(key, value) {
    if (key in appState) {
        appState[key] = value;
    } else {
        // This strict check helps catch bugs where a component tries to set
        // a state property that hasn't been formally defined.
        console.error(`Attempted to update a non-existent state key: ${key}`);
    }
}

