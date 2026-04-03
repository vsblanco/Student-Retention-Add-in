/**
 * PublishHighlightCommand – Office Script for Power Automate
 *
 * Writes a highlight command to the "SRK_Commands" sheet so that the
 * Student Retention Add-in (if active in a user's session) can pick it up
 * via a WorksheetChanged event and execute the highlight locally for
 * instant visibility.
 *
 * If no add-in is listening, the command remains in the sheet. Power Automate
 * can check after a delay and fall back to running HighlightRows.ts directly.
 *
 * @param syStudentId - The student's ID to highlight
 * @param targetSheet - Sheet name to highlight in
 * @param startCol    - Start column (name or 0-based index)
 * @param endCol      - End column (name or 0-based index)
 * @param color       - Hex highlight color (optional, omit to remove highlight)
 * @param editColumn  - Column to write text into (optional)
 * @param editText    - Text to write (optional)
 */
function main(
  workbook: ExcelScript.Workbook,
  syStudentId: string,
  targetSheet: string,
  startCol: string,
  endCol: string,
  color?: string,
  editColumn?: string,
  editText?: string
): string {
  if (!syStudentId || !targetSheet || (!startCol && startCol !== "0") || (!endCol && endCol !== "0")) {
    return "ERROR: syStudentId, targetSheet, startCol, and endCol are required.";
  }

  const COMMANDS_SHEET = "SRK_Commands";

  // Get or create the commands sheet
  let cmdSheet = workbook.getWorksheet(COMMANDS_SHEET);
  if (!cmdSheet) {
    cmdSheet = workbook.addWorksheet(COMMANDS_SHEET);
    // Set up header row
    const headerRange = cmdSheet.getRangeByIndexes(0, 0, 1, 3);
    headerRange.setValues([["Command", "Timestamp", "Status"]]);
    headerRange.getFormat().getFont().setBold(true);
    // Hide the sheet so users don't see it
    cmdSheet.setVisibility(ExcelScript.SheetVisibility.hidden);
  }

  // Build the command payload (same shape as SRK_HIGHLIGHT_STUDENT_ROW)
  const command = {
    type: "SRK_HIGHLIGHT_STUDENT_ROW",
    data: {
      syStudentId: syStudentId.trim(),
      targetSheet: targetSheet.trim(),
      startCol: startCol.trim(),
      endCol: endCol.trim(),
      color: color && color.trim() !== "" ? color.trim() : undefined,
      editColumn: editColumn && editColumn.trim() !== "" ? editColumn.trim() : undefined,
      editText: editText != null && String(editText).trim() !== "" ? editText : undefined
    }
  };

  const commandJson = JSON.stringify(command);
  const timestamp = new Date().toISOString();

  // Find the next empty row (after header)
  const usedRange = cmdSheet.getUsedRange();
  let nextRow = 1;
  if (usedRange) {
    const values = usedRange.getValues();
    nextRow = values.length;
  }

  // Write the command
  const cmdRange = cmdSheet.getRangeByIndexes(nextRow, 0, 1, 3);
  cmdRange.setValues([[commandJson, timestamp, "pending"]]);

  return `SUCCESS: Command published to ${COMMANDS_SHEET} row ${nextRow + 1} for student "${syStudentId}".`;
}
