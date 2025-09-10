// Holds the global state for the application.
// This single object makes it easier to manage and pass state between modules.

// The state object with initial values.
const appState = {
    powerAutomateUrl: null,
    students: [],
    customParameters: [],
    templates: [],
    lastFocusedElement: null,
    // The parts for the pillbox inputs are stored here
    emailParts: {
        from: [],
        subject: [],
        cc: []
    }
};

// A mutable variable to hold the Quill editor instance.
// It's kept separate from appState because it's a complex object, not simple data.
let quillInstance = null;

/**
 * Provides controlled access to the application state.
 * @returns {object} The current application state.
 */
export function getState() {
    return appState;
}

/**
 * Updates a specific part of the application state.
 * @param {string} key - The top-level key in the state object to update.
 * @param {*} value - The new value for the key.
 */
export function updateState(key, value) {
    if (appState.hasOwnProperty(key)) {
        appState[key] = value;
    } else {
        console.error(`Attempted to update a non-existent state key: ${key}`);
    }
}

/**
 * Stores the Quill editor instance.
 * @param {object} quill - The initialized Quill editor object.
 */
export function setQuill(quill) {
    quillInstance = quill;
}

/**
 * Retrieves the Quill editor instance.
 * @returns {object} The Quill editor object.
 */
export function getQuill() {
    return quillInstance;
}
