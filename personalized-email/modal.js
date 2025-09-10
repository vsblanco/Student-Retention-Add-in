/**
 * A reusable class to manage the behavior of modal dialogs.
 */
export class Modal {
    /**
     * @param {string} modalId The ID of the modal's root element.
     * @param {object} [options] Configuration options for the modal.
     * @param {string[]} [options.closeButtonIds] An array of IDs for buttons that should close the modal.
     */
    constructor(modalId, options = {}) {
        this.modalElement = document.getElementById(modalId);
        if (!this.modalElement) {
            throw new Error(`Modal element with ID "${modalId}" not found.`);
        }

        this.options = options;
        this.boundHandleKeyDown = this.handleKeyDown.bind(this);

        this.attachEventListeners();
    }

    /**
     * Attaches event listeners to the modal's close buttons and the document.
     * @private
     */
    attachEventListeners() {
        if (this.options.closeButtonIds) {
            this.options.closeButtonIds.forEach(buttonId => {
                const button = document.getElementById(buttonId);
                if (button) {
                    button.addEventListener('click', () => this.hide());
                }
            });
        }
    }

    /**
     * Shows the modal and adds a keydown listener for the 'Escape' key.
     */
    show() {
        this.modalElement.classList.remove('hidden');
        document.addEventListener('keydown', this.boundHandleKeyDown);
    }

    /**
     * Hides the modal and removes the keydown listener.
     */
    hide() {
        this.modalElement.classList.add('hidden');
        document.removeEventListener('keydown', this.boundHandleKeyDown);
    }

    /**
     * Handles the keydown event to close the modal on 'Escape'.
     * @param {KeyboardEvent} event The keyboard event.
     * @private
     */
    handleKeyDown(event) {
        if (event.key === 'Escape') {
            this.hide();
        }
    }

    /**
     * Finds and returns an element within the modal by its selector.
     * @param {string} selector The CSS selector for the element.
     * @returns {HTMLElement|null} The found element or null.
     */
    querySelector(selector) {
        return this.modalElement.querySelector(selector);
    }
}
