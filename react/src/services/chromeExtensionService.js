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

      case "SRK_IMPORT_MASTER_LIST":
        console.log("ChromeExtensionService: Import master list request received:", event.data);
        // Forward to listeners (background-service.js will handle the actual import)
        this.notifyListeners({ type: "message", data: event.data });
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
   * Send a highlight confirmation message back to the Chrome extension via window.postMessage
   * @param {string} syStudentId - The student ID (may be empty if unknown)
   * @param {"success"|"error"} status - Whether the highlight succeeded or failed
   * @param {string} message - Human-readable description of the result
   */
  sendHighlightConfirmation(syStudentId, status, message) {
    const confirmation = {
      type: "SRK_HIGHLIGHT_CONFIRMATION",
      data: {
        syStudentId: syStudentId || "",
        status,
        message,
        timestamp: Date.now()
      }
    };

    console.log(`ChromeExtensionService: Sending highlight confirmation (${status}):`, confirmation);
    this.sendMessage(confirmation);
  }

  /**
   * Handle highlighting a student row based on extension request
   * Runs in the background - no side panel required
   * @param {Object} payload - Highlight request payload
   * @param {string} payload.studentName - Student's full name
   * @param {string} payload.syStudentId - Student's SyStudentID
   * @param {number|string} payload.startCol - Starting column (0-based index OR column name)
   * @param {number|string} payload.endCol - Ending column (0-based index OR column name)
   * @param {string} payload.targetSheet - Name of the worksheet to highlight in
   * @param {string} [payload.color='yellow'] - Hex color code for highlight (optional)
   * @param {number|string} [payload.editColumn] - Column to edit (0-based index OR column name, optional)
   * @param {string} [payload.editText] - Text to set in the edit column (optional)
   */
  async handleHighlightStudentRow(payload) {
    // Validate Excel is available
    if (typeof window.Excel === "undefined") {
      const msg = "Excel API is not available. The add-in may not be running inside Excel.";
      console.warn("ChromeExtensionService:", msg);
      this.sendHighlightConfirmation(payload?.syStudentId, "error", msg);
      return;
    }

    // Validate required parameters
    if (!payload || !payload.syStudentId || !payload.targetSheet) {
      const msg = "Missing required parameters: syStudentId and targetSheet are required.";
      console.error("ChromeExtensionService:", msg, payload);
      this.sendHighlightConfirmation(payload?.syStudentId, "error", msg);
      return;
    }

    const { studentName, syStudentId, startCol, endCol, targetSheet, color = '#FFFF00', editColumn, editText } = payload;

    // Validate column parameters (can be number or string)
    if (startCol === undefined || startCol === null) {
      const msg = "startCol is required but was not provided.";
      console.error("ChromeExtensionService:", msg, { startCol });
      this.sendHighlightConfirmation(syStudentId, "error", msg);
      return;
    }

    if (endCol === undefined || endCol === null) {
      const msg = "endCol is required but was not provided.";
      console.error("ChromeExtensionService:", msg, { endCol });
      this.sendHighlightConfirmation(syStudentId, "error", msg);
      return;
    }

    // Validate types
    if (typeof startCol !== 'number' && typeof startCol !== 'string') {
      const msg = `startCol must be a number or string, but received ${typeof startCol}.`;
      console.error("ChromeExtensionService:", msg, { startCol });
      this.sendHighlightConfirmation(syStudentId, "error", msg);
      return;
    }

    if (typeof endCol !== 'number' && typeof endCol !== 'string') {
      const msg = `endCol must be a number or string, but received ${typeof endCol}.`;
      console.error("ChromeExtensionService:", msg, { endCol });
      this.sendHighlightConfirmation(syStudentId, "error", msg);
      return;
    }

    if (editColumn !== undefined && typeof editColumn !== 'number' && typeof editColumn !== 'string') {
      const msg = `editColumn must be a number or string, but received ${typeof editColumn}.`;
      console.error("ChromeExtensionService:", msg, { editColumn });
      this.sendHighlightConfirmation(syStudentId, "error", msg);
      return;
    }

    // If editColumn is provided, editText should also be provided
    if (editColumn !== undefined && editText === undefined) {
      console.warn("ChromeExtensionService: editColumn specified but editText is missing");
    }

    try {
      await Excel.run(async (context) => {
        // Helper function to normalize date formats (e.g., "01/11/2026" <-> "1/11/2026" <-> "LDA 01-11-2026")
        const normalizeDateFormat = (dateStr) => {
          // Check if the string looks like a date format (contains / or -)
          if (!dateStr || typeof dateStr !== 'string') {
            return [dateStr]; // Not a string, return as-is
          }

          // Extract any prefix (e.g., "LDA " from "LDA 01-11-2026")
          const datePattern = /^(.*?)(\d{1,2}[/-]\d{1,2}[/-]\d{4})$/;
          const match = dateStr.match(datePattern);

          if (!match) {
            return [dateStr]; // Not a recognizable date format, return as-is
          }

          const prefix = match[1]; // e.g., "LDA " or ""
          const datePart = match[2]; // e.g., "01-11-2026"

          // Detect separator (slash or dash)
          let separator = null;
          if (datePart.includes('/')) {
            separator = '/';
          } else if (datePart.includes('-')) {
            separator = '-';
          } else {
            return [dateStr]; // Not a date format, return as-is
          }

          const parts = datePart.split(separator);
          if (parts.length !== 3) {
            return [dateStr]; // Not a standard date format
          }

          const [month, day, year] = parts;

          // Generate variations with and without leading zeros, with both separators
          const variations = new Set();

          // Add original
          variations.add(dateStr);

          // For both separators (/ and -)
          const separators = ['/', '-'];
          separators.forEach(sep => {
            // Version with leading zeros
            variations.add(`${prefix}${month.padStart(2, '0')}${sep}${day.padStart(2, '0')}${sep}${year}`);

            // Version without leading zeros
            variations.add(`${prefix}${parseInt(month, 10)}${sep}${parseInt(day, 10)}${sep}${year}`);

            // Variations with only month or only day having leading zeros
            variations.add(`${prefix}${month.padStart(2, '0')}${sep}${parseInt(day, 10)}${sep}${year}`);
            variations.add(`${prefix}${parseInt(month, 10)}${sep}${day.padStart(2, '0')}${sep}${year}`);
          });

          return Array.from(variations);
        };

        // Try to get the target worksheet with resilient matching
        let worksheet = context.workbook.worksheets.getItemOrNullObject(targetSheet);
        worksheet.load("isNullObject");
        await context.sync();

        // If exact match fails, try normalized date variations
        if (worksheet.isNullObject) {
          const variations = normalizeDateFormat(targetSheet);

          // If we have variations, load all sheets and try to find a match
          if (variations.length > 1) {
            const sheets = context.workbook.worksheets;
            sheets.load("items/name");
            await context.sync();

            // Try each variation against all sheet names
            let foundSheetName = null;
            for (const variation of variations) {
              const matchingSheet = sheets.items.find(sheet => sheet.name === variation);
              if (matchingSheet) {
                foundSheetName = matchingSheet.name;
                break;
              }
            }

            if (foundSheetName) {
              console.log(`ChromeExtensionService: Found sheet "${foundSheetName}" using date format normalization (requested: "${targetSheet}")`);
              worksheet = context.workbook.worksheets.getItem(foundSheetName);
            }
          }
        }

        // Final check if sheet was found
        worksheet.load("isNullObject");
        await context.sync();

        if (worksheet.isNullObject) {
          const msg = `Sheet "${targetSheet}" not found in the workbook. Verify the sheet name is correct (sheet names are case-sensitive).`;
          console.error("ChromeExtensionService:", msg);
          this.sendHighlightConfirmation(syStudentId, "error", msg);
          return;
        }

        // Load the used range to search for student
        const usedRange = worksheet.getUsedRange();
        usedRange.load(["values", "rowCount", "columnCount", "rowIndex"]);
        await context.sync();

        const values = usedRange.values;
        const headers = values[0]; // First row is headers

        // Helper function to resolve column name to index
        // Returns { index, error } where error is a message string if resolution failed
        const resolveColumnIndex = (column, paramName) => {
          // If it's already a number, validate and return it
          if (typeof column === 'number') {
            if (column < 0) {
              const msg = `${paramName} must be >= 0, but received ${column}.`;
              console.error("ChromeExtensionService:", msg, { [paramName]: column });
              return { index: -1, error: msg };
            }
            if (column >= headers.length) {
              const msg = `${paramName} index (${column}) exceeds the number of columns in the sheet (${headers.length}).`;
              console.error("ChromeExtensionService:", msg, { [paramName]: column });
              return { index: -1, error: msg };
            }
            return { index: column, error: null };
          }

          // It's a string, find the column by name
          const columnName = String(column).trim();
          for (let col = 0; col < headers.length; col++) {
            if (String(headers[col]).trim().toLowerCase() === columnName.toLowerCase()) {
              return { index: col, error: null };
            }
          }

          // Column name not found
          const msg = `Column "${columnName}" not found in sheet "${targetSheet}". Check that the column header exists in row 1.`;
          console.error("ChromeExtensionService:", msg, { [paramName]: column });
          return { index: -1, error: msg };
        };

        // Resolve startCol, endCol, and editColumn to indices
        const startColResult = resolveColumnIndex(startCol, 'startCol');
        const endColResult = resolveColumnIndex(endCol, 'endCol');

        if (startColResult.error || endColResult.error) {
          const msg = startColResult.error || endColResult.error;
          this.sendHighlightConfirmation(syStudentId, "error", msg);
          return;
        }

        const startColIndex = startColResult.index;
        const endColIndex = endColResult.index;

        if (startColIndex > endColIndex) {
          const msg = `startCol (${startCol}, index ${startColIndex}) must be less than or equal to endCol (${endCol}, index ${endColIndex}).`;
          console.error("ChromeExtensionService:", msg);
          this.sendHighlightConfirmation(syStudentId, "error", msg);
          return;
        }

        let editColumnIndex = -1;
        if (editColumn !== undefined) {
          const editColResult = resolveColumnIndex(editColumn, 'editColumn');
          if (editColResult.error) {
            this.sendHighlightConfirmation(syStudentId, "error", editColResult.error);
            return;
          }
          editColumnIndex = editColResult.index;
        }

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
          const msg = `Could not find a Student ID column in sheet "${targetSheet}". Expected one of: ${idAliases.join(', ')}.`;
          console.error("ChromeExtensionService:", msg);
          this.sendHighlightConfirmation(syStudentId, "error", msg);
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
          const msg = `Student with ID "${syStudentId}" not found in sheet "${targetSheet}". Verify the student ID exists in the sheet.`;
          console.warn("ChromeExtensionService:", msg);
          this.sendHighlightConfirmation(syStudentId, "error", msg);
          return;
        }

        // Calculate the actual row index in the worksheet (accounting for used range offset)
        const actualRowIndex = usedRange.rowIndex + targetRowIndex;

        // Calculate column count to highlight
        const columnCount = endColIndex - startColIndex + 1;

        // Get the range to highlight
        const highlightRange = worksheet.getRangeByIndexes(
          actualRowIndex,
          startColIndex,
          1, // Just one row
          columnCount
        );

        // Apply the highlight color
        highlightRange.format.fill.color = color;
        await context.sync();

        // Edit cell if editColumn and editText are provided
        if (editColumn !== undefined && editText !== undefined && editColumnIndex !== -1) {
          const editCell = worksheet.getRangeByIndexes(
            actualRowIndex,
            editColumnIndex,
            1, // One row
            1  // One column
          );
          editCell.values = [[editText]];
          await context.sync();

          console.log(`ChromeExtensionService: Successfully edited cell at column "${editColumn}" (index ${editColumnIndex}) with text "${editText}"`);
        }

        console.log(`ChromeExtensionService: Successfully highlighted student "${studentName}" (ID: ${syStudentId}) in "${targetSheet}" from column "${startCol}" (index ${startColIndex}) to "${endCol}" (index ${endColIndex})`);

        // Send confirmation back to Chrome extension
        this.sendHighlightConfirmation(syStudentId, "success", "Row highlighted successfully");

        // Notify listeners of successful highlight
        this.notifyListeners({
          type: "highlight_complete",
          data: {
            studentName,
            syStudentId,
            targetSheet,
            startCol,
            endCol,
            startColIndex,
            endColIndex,
            color,
            editColumn,
            editColumnIndex,
            editText,
            timestamp: new Date().toISOString()
          }
        });
      });
    } catch (error) {
      console.error("ChromeExtensionService: Error highlighting student row:", error);

      // Determine a helpful error message based on common failure scenarios
      let errorMessage = error.message || "An unknown error occurred while highlighting the row.";
      if (error.code === "InvalidReference") {
        errorMessage = "Invalid cell reference. The worksheet structure may have changed.";
      } else if (error.code === "GeneralException") {
        errorMessage = "Excel encountered an error. The workbook may be protected, read-only, or in use by another process.";
      } else if (error.code === "AccessDenied") {
        errorMessage = "Access denied. The workbook or sheet may be protected.";
      } else if (error.message && error.message.includes("network")) {
        errorMessage = "A network error occurred. Check your connection to Excel Online.";
      }

      // Send error confirmation back to Chrome extension
      this.sendHighlightConfirmation(syStudentId, "error", errorMessage);

      // Notify listeners of error
      this.notifyListeners({
        type: "highlight_error",
        data: {
          studentName,
          syStudentId,
          editColumn,
          editText,
          error: errorMessage,
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
    // If there's a parent window (Office Add-in running in iframe), send only to parent
    // This prevents duplicate messages to the Chrome extension
    if (window.parent && window.parent !== window) {
      console.log(`📤 ChromeExtensionService: Sending message (${message.type}) to parent window:`, message);
      window.parent.postMessage(message, "*");
    } else {
      // No parent window, send to self
      console.log(`📤 ChromeExtensionService: Sending message (${message.type}) to window:`, message);
      window.postMessage(message, "*");
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
