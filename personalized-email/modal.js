/**
 * A reusable class to manage modal dialogs.
 */
export class Modal {
    /**
     * Creates an instance of a Modal.
     * @param {string} modalId The ID of the main modal element.
     * @param {object} [options] Configuration options for the modal.
     * @param {string} [options.closeButtonId] The ID of the button that closes the modal.
     */
    constructor(modalId, options = {}) {
        this.modalElement = document.getElementById(modalId);
        this.options = options;

        if (!this.modalElement) {
            console.error(`Modal element with ID "${modalId}" not found.`);
            return;
        }

        this.initCloseListeners();
    }

    /**
     * Initializes the event listeners for closing the modal.
     */
    initCloseListeners() {
        if (this.options.closeButtonId) {
            // FIX: Use document.getElementById for a more robust lookup.
            const closeButton = document.getElementById(this.options.closeButtonId);
            if (closeButton) {
                closeButton.onclick = () => this.hide();
            } else {
                console.warn(`Close button with ID "${this.options.closeButtonId}" not found in the document.`);
            }
        }
    }

    /**
     * Shows the modal.
     */
    show() {
        if (this.modalElement) {
            this.modalElement.classList.remove('hidden');
        }
    }

    /**
     * Hides the modal.
     */
    hide() {
        if (this.modalElement) {
            this.modalElement.classList.add('hidden');
        }
    }
}

