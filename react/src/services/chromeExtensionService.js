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

      case "SRK_HIGHLIGHT_STUDENT_ROW":
        console.log("ChromeExtensionService: Highlight student row request received:", event.data);
        this.handleHighlightStudentRow(event.data.data);
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
   * Handle highlighting a student row based on extension request
   * Runs in the background - no side panel required
   * @param {Object} payload - Highlight request payload
   * @param {string} payload.studentName - Student's full name
   * @param {string} payload.syStudentId - Student's SyStudentID
   * @param {number} payload.startCol - Starting column index (0-based)
   * @param {number} payload.endCol - Ending column index (0-based)
   * @param {string} payload.targetSheet - Name of the worksheet to highlight in
   * @param {string} [payload.color='yellow'] - Hex color code for highlight (optional)
   */
  async handleHighlightStudentRow(payload) {
    // Validate Excel is available
    if (typeof window.Excel === "undefined") {
      console.warn("ChromeExtensionService: Excel API not available");
      return;
    }

    // Validate required parameters
    if (!payload || !payload.syStudentId || !payload.targetSheet) {
      console.error("ChromeExtensionService: Missing required parameters for highlight", payload);
      return;
    }

    const { studentName, syStudentId, startCol, endCol, targetSheet, color = '#FFFF00' } = payload;

    // Validate column indices
    if (typeof startCol !== 'number' || typeof endCol !== 'number') {
      console.error("ChromeExtensionService: Invalid column indices", { startCol, endCol });
      return;
    }

    if (startCol < 0 || endCol < 0 || startCol > endCol) {
      console.error("ChromeExtensionService: Invalid column range", { startCol, endCol });
      return;
    }

    try {
      await Excel.run(async (context) => {
        // Get the target worksheet
        const worksheet = context.workbook.worksheets.getItemOrNullObject(targetSheet);
        worksheet.load("isNullObject");
        await context.sync();

        if (worksheet.isNullObject) {
          console.error(`ChromeExtensionService: Sheet "${targetSheet}" not found`);
          return;
        }

        // Load the used range to search for student
        const usedRange = worksheet.getUsedRange();
        usedRange.load(["values", "rowCount", "columnCount"]);
        await context.sync();

        const values = usedRange.values;
        const headers = values[0]; // First row is headers

        // Find the ID column index
        let idColumnIndex = -1;
        const idAliases = ['Student ID', 'SyStudentID', 'Student identifier', 'ID'];

        for (let col = 0; col < headers.length; col++) {
          if (idAliases.some(alias =>
            String(headers[col]).trim().toLowerCase() === alias.toLowerCase()
          )) {
            idColumnIndex = col;
            break;
          }
        }

        if (idColumnIndex === -1) {
          console.error("ChromeExtensionService: Could not find Student ID column in sheet");
          return;
        }

        // Find the row with matching syStudentId
        let targetRowIndex = -1;
        for (let row = 1; row < values.length; row++) { // Start from 1 to skip headers
          const cellValue = String(values[row][idColumnIndex]).trim();
          if (cellValue === String(syStudentId).trim()) {
            targetRowIndex = row;
            break;
          }
        }

        if (targetRowIndex === -1) {
          console.warn(`ChromeExtensionService: Student with ID "${syStudentId}" not found in sheet "${targetSheet}"`);
          return;
        }

        // Calculate the actual row index in the worksheet (accounting for used range offset)
        const actualRowIndex = usedRange.rowIndex + targetRowIndex;

        // Calculate column count to highlight
        const columnCount = endCol - startCol + 1;

        // Get the range to highlight
        const highlightRange = worksheet.getRangeByIndexes(
          actualRowIndex,
          startCol,
          1, // Just one row
          columnCount
        );

        // Apply the highlight color
        highlightRange.format.fill.color = color;
        await context.sync();

        console.log(`ChromeExtensionService: Successfully highlighted student "${studentName}" (ID: ${syStudentId}) in "${targetSheet}" from column ${startCol} to ${endCol}`);

        // Notify listeners of successful highlight
        this.notifyListeners({
          type: "highlight_complete",
          data: {
            studentName,
            syStudentId,
            targetSheet,
            startCol,
            endCol,
            color,
            timestamp: new Date().toISOString()
          }
        });
      });
    } catch (error) {
      console.error("ChromeExtensionService: Error highlighting student row:", error);

      // Notify listeners of error
      this.notifyListeners({
        type: "highlight_error",
        data: {
          studentName,
          syStudentId,
          error: error.message,
          timestamp: new Date().toISOString()
        }
      });
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
   * Send selected student data to the Chrome extension
   * @param {Array|Object} students - Single student object or array of students
   */
  sendSelectedStudents(students) {
    // Normalize to array format
    const studentArray = Array.isArray(students) ? students : [students];

    // Extract only the required fields
    const payload = studentArray.map(student => ({
      name: student.StudentName || student.name || "",
      syStudentId: student.ID || student.syStudentId || "",
      phone: student.Phone || student.phone || "",
      otherPhone: student.OtherPhone || student.otherPhone || ""
    }));

    const message = {
      type: "SRK_SELECTED_STUDENTS",
      data: {
        students: payload,
        count: payload.length,
        timestamp: new Date().toISOString()
      }
    };

    console.log("ChromeExtensionService: Sending selected students:", message);
    this.sendMessage(message);
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
