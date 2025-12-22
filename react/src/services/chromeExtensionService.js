/**
 * Chrome Extension Communication Service
 *
 * Centralized service for managing communication with the Student Retention Kit Chrome Extension.
 * Handles extension detection, keep-alive pings, and message relay between the add-in and extension.
 */

class ChromeExtensionService {
  constructor() {
    // Connection state
    this.isExtensionInstalled = false;
    this.isChecking = false;

    // Intervals and timeouts
    this.pingInterval = null;
    this.keepAliveInterval = null;

    // Callbacks for extension state changes
    this.listeners = new Set();

    // Message handler bound to this instance
    this.handleMessage = this.handleMessage.bind(this);

    // Setup message listener
    this.setupMessageListener();
  }

  /**
   * Setup the global message listener for extension responses
   */
  setupMessageListener() {
    window.addEventListener("message", this.handleMessage);
  }

  /**
   * Handle incoming messages from the Chrome extension
   */
  handleMessage(event) {
    if (!event.data || !event.data.type) return;

    switch (event.data.type) {
      case "SRK_EXTENSION_INSTALLED":
        console.log("ChromeExtensionService: Extension detected!");
        this.handleExtensionDetected();
        break;

      // Add more message types here as needed
      default:
        // Forward unknown messages to listeners
        this.notifyListeners({ type: "message", data: event.data });
    }
  }

  /**
   * Handle when extension is detected
   */
  handleExtensionDetected() {
    if (!this.isExtensionInstalled) {
      this.isExtensionInstalled = true;
      this.stopPinging();
      this.notifyListeners({ type: "installed", installed: true });
    }
  }

  /**
   * Send a ping to check if extension is installed
   * Sends to both current window and parent window
   */
  sendPing() {
    const message = { type: "SRK_CHECK_EXTENSION" };

    // Send to self
    window.postMessage(message, "*");

    // Send to parent if in iframe
    if (window.parent && window.parent !== window) {
      window.parent.postMessage(message, "*");
    }
  }

  /**
   * Start pinging to detect the extension
   * @param {number} interval - Milliseconds between pings (default: 2000)
   */
  startPinging(interval = 2000) {
    if (this.isChecking) return;

    console.log("ChromeExtensionService: Starting extension detection...");
    this.isChecking = true;

    // Send initial ping
    this.sendPing();

    // Setup interval for periodic pings
    this.pingInterval = setInterval(() => {
      if (!this.isExtensionInstalled) {
        this.sendPing();
      } else {
        this.stopPinging();
      }
    }, interval);
  }

  /**
   * Stop pinging for extension detection
   */
  stopPinging() {
    if (this.pingInterval) {
      console.log("ChromeExtensionService: Stopping extension detection.");
      clearInterval(this.pingInterval);
      this.pingInterval = null;
      this.isChecking = false;
    }
  }

  /**
   * Start keep-alive heartbeat for Chrome extension context
   * This prevents the extension from going dormant
   * @param {number} interval - Milliseconds between pings (default: 20000)
   */
  startKeepAlive(interval = 20000) {
    if (this.keepAliveInterval) return;

    console.log("ChromeExtensionService: Starting keep-alive heartbeat...");

    this.keepAliveInterval = setInterval(() => {
      if (typeof chrome !== 'undefined' && chrome.runtime && chrome.runtime.sendMessage) {
        chrome.runtime.sendMessage({ type: 'keep_alive_ping' }, (response) => {
          if (chrome.runtime.lastError) {
            // Ignore errors - extension might not be installed
          }
        });
      }
    }, interval);
  }

  /**
   * Stop keep-alive heartbeat
   */
  stopKeepAlive() {
    if (this.keepAliveInterval) {
      console.log("ChromeExtensionService: Stopping keep-alive heartbeat.");
      clearInterval(this.keepAliveInterval);
      this.keepAliveInterval = null;
    }
  }

  /**
   * Add a listener for extension events
   * @param {Function} callback - Function to call when events occur
   * @returns {Function} Cleanup function to remove the listener
   */
  addListener(callback) {
    this.listeners.add(callback);

    // Return cleanup function
    return () => {
      this.listeners.delete(callback);
    };
  }

  /**
   * Notify all listeners of an event
   * @param {Object} event - Event object to send to listeners
   */
  notifyListeners(event) {
    this.listeners.forEach(listener => {
      try {
        listener(event);
      } catch (error) {
        console.error("ChromeExtensionService: Listener error:", error);
      }
    });
  }

  /**
   * Send a custom message to the extension
   * @param {Object} message - Message object to send
   */
  sendMessage(message) {
    window.postMessage(message, "*");

    if (window.parent && window.parent !== window) {
      window.parent.postMessage(message, "*");
    }
  }

  /**
   * Send a message using chrome.runtime API (if available)
   * @param {Object} message - Message to send
   * @param {Function} callback - Optional callback for response
   */
  sendChromeMessage(message, callback) {
    if (typeof chrome !== 'undefined' && chrome.runtime && chrome.runtime.sendMessage) {
      chrome.runtime.sendMessage(message, (response) => {
        if (chrome.runtime.lastError) {
          console.warn("ChromeExtensionService: Chrome message error:", chrome.runtime.lastError);
        }
        if (callback) callback(response);
      });
    } else {
      console.warn("ChromeExtensionService: chrome.runtime.sendMessage not available");
    }
  }

  /**
   * Get current extension installation status
   * @returns {boolean} Whether the extension is detected as installed
   */
  getInstallationStatus() {
    return this.isExtensionInstalled;
  }

  /**
   * Reset the service state (useful for testing or reinitializing)
   */
  reset() {
    this.stopPinging();
    this.stopKeepAlive();
    this.isExtensionInstalled = false;
    this.isChecking = false;
  }

  /**
   * Cleanup all intervals and listeners
   */
  cleanup() {
    this.stopPinging();
    this.stopKeepAlive();
    window.removeEventListener("message", this.handleMessage);
    this.listeners.clear();
  }
}

// Create and export a singleton instance
const chromeExtensionService = new ChromeExtensionService();

export default chromeExtensionService;
