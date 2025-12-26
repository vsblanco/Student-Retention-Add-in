/*
 * add-user-modal.js
 * Version: 1.0.0
 *
 * Manages the "Add User" modal dialog functionality.
 * Handles user input validation, duplicate checking, and user list updates.
 */
function initializeAddUserModal(settings, saveSettingsCallback) {
    // --- Get DOM Elements ---
    const addUserModal = document.getElementById(CONSTANTS.DOM.IDS.addUserModal);
    const addUserButton = document.getElementById(CONSTANTS.DOM.IDS.addUserButton);
    const cancelAddUserButton = document.getElementById(CONSTANTS.DOM.IDS.cancelAddUserButton);
    const saveNewUserButton = document.getElementById(CONSTANTS.DOM.IDS.saveNewUserButton);
    const firstNameInput = document.getElementById(CONSTANTS.DOM.IDS.newUserFirstName);
    const lastNameInput = document.getElementById(CONSTANTS.DOM.IDS.newUserLastName);
    const addUserStatus = document.getElementById(CONSTANTS.DOM.IDS.addUserStatus);

    // --- Modal Visibility Functions ---

    /**
     * Resets the modal form and makes it visible.
     */
    const showModal = () => {
        // Clear previous status messages and input values
        addUserStatus.textContent = '';
        addUserStatus.className = CONSTANTS.DOM.CLASSES.statusMessage;
        firstNameInput.value = '';
        lastNameInput.value = '';
        
        // Display the modal and focus the first input field
        addUserModal.style.display = 'flex';
        firstNameInput.focus();
    };

    /**
     * Hides the modal from view.
     */
    const hideModal = () => {
        addUserModal.style.display = 'none';
    };

    // --- Event Handlers ---

    /**
     * Handles the click event for the "Save New User" button.
     * Validates input, checks for duplicates, adds the new user, and saves settings.
     */
    const handleSaveUser = () => {
        const firstName = firstNameInput.value.trim();
        const lastName = lastNameInput.value.trim();

        // Validate that both names are provided
        if (!firstName || !lastName) {
            addUserStatus.textContent = "Both names are required.";
            addUserStatus.className = `${CONSTANTS.DOM.CLASSES.statusMessage} ${CONSTANTS.DOM.CLASSES.statusError} ${CONSTANTS.DOM.CLASSES.visible}`;
            return;
        }

        const formattedName = `${firstName} ${lastName}`;

        // Check if the user already exists in the list
        if (settings.userProfile.userList.includes(formattedName)) {
            addUserStatus.textContent = `User "${formattedName}" already exists.`;
            addUserStatus.className = `${CONSTANTS.DOM.CLASSES.statusMessage} ${CONSTANTS.DOM.CLASSES.statusError} ${CONSTANTS.DOM.CLASSES.visible}`;
            return;
        }

        // Add the new user and trigger the save callback
        settings.userProfile.userList.push(formattedName);
        if (typeof saveSettingsCallback === 'function') {
            saveSettingsCallback(); // This will save and reload the UI from settings.js
        }
        hideModal();
    };

    /**
     * Handles the keydown event on the input fields to allow saving with "Enter".
     * @param {KeyboardEvent} event - The keyboard event object.
     */
    const handleInputKeyDown = (event) => {
        if (event.key === 'Enter') {
            handleSaveUser();
        } else if (event.key === 'Escape') {
            hideModal();
        }
    };

    // --- Attach Event Listeners ---
    addUserButton.onclick = showModal;
    cancelAddUserButton.onclick = hideModal;
    saveNewUserButton.onclick = handleSaveUser;

    // Add keydown listeners to both input fields
    firstNameInput.addEventListener('keydown', handleInputKeyDown);
    lastNameInput.addEventListener('keydown', handleInputKeyDown);

    // Also hide the modal if the user clicks the dark overlay background
    addUserModal.addEventListener('click', (event) => {
        // We check if the clicked element is the modal overlay itself, not a child.
        if (event.target === addUserModal) {
            hideModal();
        }
    });
}