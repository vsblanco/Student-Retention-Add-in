/*
 * power-automate-poller.js
 *
 * Polls the workbook's custom document property "SRK_Command" for highlight
 * commands written by Power Automate Office Scripts. Custom properties sync
 * via co-authoring much faster than cell values or formatting, enabling
 * near real-time highlights without a page refresh.
 *
 * NOTE: This is the Power Automate path, not the Chrome extension path.
 * Chrome extension messaging lives in ./chrome-extension-messaging.js.
 *
 * Claim mechanism for multi-user:
 *   Each add-in instance has a unique session ID. When a pending command
 *   is detected, the instance writes its ID to "SRK_CommandClaim". After
 *   a brief delay, it re-reads the claim — if its ID is still there, it
 *   won the claim and executes. Otherwise, another instance handles it.
 *
 * Flow:
 *   1. Power Automate runs HighlightRows.ts which applies direct formatting
 *      AND writes the command to "SRK_Command" custom property
 *   2. Co-authoring syncs the property to all active sessions (~1-5 seconds)
 *   3. First add-in to claim the command executes the highlight locally
 *      (instant visibility), then clears the property
 */

/**
 * Highlights a student's row directly via the Excel API.
 * Standalone version — no Chrome extension messaging, no postMessage,
 * no listener notifications. Used by the command property poller.
 *
 * @param {Object} payload - Same shape as SRK_HIGHLIGHT_STUDENT_ROW
 * @param {string} payload.syStudentId - Student ID to find
 * @param {string} payload.targetSheet - Sheet name
 * @param {string|number} payload.startCol - Start column (name or index)
 * @param {string|number} payload.endCol - End column (name or index)
 * @param {string} [payload.color='#FFFF00'] - Highlight color
 * @param {string|number} [payload.editColumn] - Column to write text into
 * @param {string} [payload.editText] - Text to write
 */
async function highlightStudentRow(payload) {
    if (!payload || !payload.syStudentId || !payload.targetSheet) {
        console.error("highlightStudentRow: Missing required parameters (syStudentId, targetSheet)");
        return;
    }

    const { syStudentId, targetSheet, startCol, endCol, color = '#FFFF00', editColumn, editText } = payload;

    if (startCol === undefined || startCol === null || endCol === undefined || endCol === null) {
        console.error("highlightStudentRow: startCol and endCol are required");
        return;
    }

    await Excel.run(async (context) => {
        // Find the sheet
        const sheets = context.workbook.worksheets;
        sheets.load("items/name");
        await context.sync();

        let worksheet = sheets.items.find(s => s.name === targetSheet);
        if (!worksheet) {
            console.error(`highlightStudentRow: Sheet "${targetSheet}" not found`);
            return;
        }
        worksheet = context.workbook.worksheets.getItem(targetSheet);

        // Load used range
        const usedRange = worksheet.getUsedRange();
        usedRange.load(["values", "rowCount", "columnCount", "rowIndex"]);
        await context.sync();

        const values = usedRange.values;
        const headers = values[0];

        // Resolve column name to index
        const resolveCol = (col) => {
            if (typeof col === 'number') return col >= 0 && col < headers.length ? col : -1;
            const colName = String(col).trim().toLowerCase();
            for (let c = 0; c < headers.length; c++) {
                if (String(headers[c]).trim().toLowerCase() === colName) return c;
            }
            return -1;
        };

        const startColIndex = resolveCol(startCol);
        const endColIndex = resolveCol(endCol);
        if (startColIndex === -1 || endColIndex === -1) {
            console.error(`highlightStudentRow: Could not resolve columns - startCol: ${startCol} (${startColIndex}), endCol: ${endCol} (${endColIndex})`);
            return;
        }

        const colStart = Math.min(startColIndex, endColIndex);
        const colEnd = Math.max(startColIndex, endColIndex);

        // Find Student ID column
        const idAliases = ['student id', 'systudentid', 'student identifier', 'id'];
        let idColIndex = -1;
        for (let c = 0; c < headers.length; c++) {
            if (idAliases.includes(String(headers[c]).trim().toLowerCase())) {
                idColIndex = c;
                break;
            }
        }
        if (idColIndex === -1) {
            console.error(`highlightStudentRow: No Student ID column found in sheet "${targetSheet}"`);
            return;
        }

        // Find the student's row
        let targetRowIndex = -1;
        for (let row = 1; row < values.length; row++) {
            if (String(values[row][idColIndex]).trim() === String(syStudentId).trim()) {
                targetRowIndex = row;
                break;
            }
        }
        if (targetRowIndex === -1) {
            console.error(`highlightStudentRow: Student ID "${syStudentId}" not found in sheet "${targetSheet}"`);
            return;
        }

        // Apply highlight
        const actualRowIndex = usedRange.rowIndex + targetRowIndex;
        const columnCount = colEnd - colStart + 1;
        const highlightRange = worksheet.getRangeByIndexes(actualRowIndex, colStart, 1, columnCount);

        // Check if highlight color needs to change
        highlightRange.format.fill.load("color");
        await context.sync();
        const currentColor = highlightRange.format.fill.color;
        const needsHighlight = !currentColor || currentColor.toLowerCase() !== color.toLowerCase();

        if (needsHighlight) {
            highlightRange.format.fill.color = color;
            await context.sync();
        }

        // Edit cell if requested and text is different
        if (editColumn !== undefined && editText !== undefined) {
            const editColIndex = resolveCol(editColumn);
            if (editColIndex !== -1) {
                const editCell = worksheet.getRangeByIndexes(actualRowIndex, editColIndex, 1, 1);
                editCell.load("values");
                await context.sync();
                const currentText = String(editCell.values[0][0] ?? "").trim();
                const newText = String(editText).trim();
                if (currentText !== newText) {
                    editCell.values = [[editText]];
                    await context.sync();
                    console.log(`highlightStudentRow: Updated text in "${editColumn}" to "${editText}"`);
                }
            }
        }

        if (needsHighlight) {
            console.log(`highlightStudentRow: Highlighted student "${syStudentId}" on "${targetSheet}" (row ${targetRowIndex + 1}, columns ${colStart}-${colEnd})`);
        } else {
            console.log(`highlightStudentRow: Student "${syStudentId}" already highlighted, no color change needed`);
        }
    });
}

const SRK_COMMAND_PROPERTY = "SRK_Command";
const SRK_CLAIM_PROPERTY = "SRK_CommandClaim";
const SRK_POLL_INTERVAL = 5000; // 5 seconds
const SRK_CLAIM_WAIT = 2000; // 2 seconds to wait before checking claim
const SRK_SESSION_ID = `session_${Date.now()}_${Math.random().toString(36).substring(2, 8)}`;
let lastProcessedTimestamp = null;
let commandPollerActive = false;
let isProcessingCommand = false;

export function startCommandPropertyPoller() {
    if (commandPollerActive) {
        console.log("CommandPropertyPoller: Already running, skipping duplicate start");
        return;
    }
    commandPollerActive = true;
    console.log(`CommandPropertyPoller: Session ID: ${SRK_SESSION_ID}`);
    console.log(`CommandPropertyPoller: Polling "${SRK_COMMAND_PROPERTY}" every ${SRK_POLL_INTERVAL / 1000}s`);

    setInterval(async () => {
        if (isProcessingCommand) return; // Don't overlap with an in-progress claim

        try {
            await Excel.run(async (context) => {
                const props = context.workbook.properties.custom;
                props.load("key, value");
                await context.sync();

                const srkProp = props.items.find(p => p.key === SRK_COMMAND_PROPERTY);
                if (!srkProp) return;

                const rawValue = srkProp.value;
                if (!rawValue || rawValue === "CLEAR") return;

                let command;
                try {
                    command = JSON.parse(rawValue);
                } catch (e) {
                    console.warn("CommandPropertyPoller: Invalid JSON, clearing:", rawValue);
                    context.workbook.properties.custom.add(SRK_COMMAND_PROPERTY, "CLEAR");
                    await context.sync();
                    return;
                }

                // Skip if already processed this command
                if (command.timestamp && command.timestamp === lastProcessedTimestamp) {
                    return;
                }

                console.log(`CommandPropertyPoller: New command detected (timestamp: ${command.timestamp})`);

                // Try to claim the command
                isProcessingCommand = true;
                context.workbook.properties.custom.add(SRK_CLAIM_PROPERTY, SRK_SESSION_ID);
                await context.sync();
            });

            if (!isProcessingCommand) return;

            // Wait briefly to let other sessions write their claim
            await new Promise(resolve => setTimeout(resolve, SRK_CLAIM_WAIT));

            // Re-read to see who won the claim
            await Excel.run(async (context) => {
                const props = context.workbook.properties.custom;
                props.load("key, value");
                await context.sync();

                const claimProp = props.items.find(p => p.key === SRK_CLAIM_PROPERTY);
                const claimValue = claimProp ? claimProp.value : null;

                // Re-read the command (it might have been cleared by the winner)
                const cmdProp = props.items.find(p => p.key === SRK_COMMAND_PROPERTY);
                const rawValue = cmdProp ? cmdProp.value : null;

                if (!rawValue || rawValue === "CLEAR") {
                    // Another session already processed and cleared it
                    console.log("CommandPropertyPoller: Command already handled by another session");
                    isProcessingCommand = false;
                    return;
                }

                let command;
                try {
                    command = JSON.parse(rawValue);
                } catch (e) {
                    isProcessingCommand = false;
                    return;
                }

                if (claimValue !== SRK_SESSION_ID) {
                    // Another session won the claim
                    console.log(`CommandPropertyPoller: Claim lost to ${claimValue}, skipping`);
                    lastProcessedTimestamp = command.timestamp;
                    isProcessingCommand = false;
                    return;
                }

                // We won the claim — execute the highlight
                console.log("CommandPropertyPoller: Claim won! Executing highlight...");

                if (command.type === "SRK_HIGHLIGHT_STUDENT_ROW" && command.data) {
                    try {
                        await highlightStudentRow(command.data);
                        console.log(`CommandPropertyPoller: Highlighted student "${command.data.syStudentId}" successfully`);
                    } catch (highlightError) {
                        console.error("CommandPropertyPoller: Highlight failed:", highlightError.message);
                    }
                } else {
                    console.warn("CommandPropertyPoller: Unknown command type:", command.type);
                }

                // Mark as processed and clear for next command
                lastProcessedTimestamp = command.timestamp;
                context.workbook.properties.custom.add(SRK_COMMAND_PROPERTY, "CLEAR");
                context.workbook.properties.custom.add(SRK_CLAIM_PROPERTY, "CLEAR");
                await context.sync();
                console.log("CommandPropertyPoller: Command completed, property cleared");
                isProcessingCommand = false;
            });
        } catch (error) {
            isProcessingCommand = false;
            if (error.code !== "GeneralException") {
                console.error("CommandPropertyPoller: Error:", error.message);
            }
        }
    }, SRK_POLL_INTERVAL);
}
