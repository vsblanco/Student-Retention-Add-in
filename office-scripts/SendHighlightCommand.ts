/**
 * SendHighlightCommand – Office Script for Power Automate
 *
 * Writes a highlight command to the "SRK_Command" custom document property.
 * The Office Add-in (if active in a user's session) polls this property
 * and executes the highlight locally for instant visibility — no page
 * refresh needed.
 *
 * Custom document properties sync via co-authoring within seconds,
 * unlike cell values or formatting which can be delayed.
 *
 * Uses the same payload shape as the Chrome Extension's
 * SRK_HIGHLIGHT_STUDENT_ROW message.
 *
 * @param syStudentId - The student's ID to highlight
 * @param targetSheet - Sheet name to highlight in
 * @param startCol    - Start column (name or 0-based index)
 * @param endCol      - End column (name or 0-based index)
 * @param color       - Hex highlight color (optional, omit to remove)
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
  // Validate required params
  if (!syStudentId || syStudentId.trim() === "") {
    return "ERROR: syStudentId is required.";
  }
  if (!targetSheet || targetSheet.trim() === "") {
    return "ERROR: targetSheet is required.";
  }
  if (!startCol && startCol !== "0") {
    return "ERROR: startCol is required.";
  }
  if (!endCol && endCol !== "0") {
    return "ERROR: endCol is required.";
  }

  // Build the command (same shape as SRK_HIGHLIGHT_STUDENT_ROW)
  const command: Record<string, unknown> = {
    type: "SRK_HIGHLIGHT_STUDENT_ROW",
    data: {
      syStudentId: syStudentId.trim(),
      targetSheet: targetSheet.trim(),
      startCol: startCol.trim(),
      endCol: endCol.trim(),
      color: color && color.trim() !== "" ? color.trim() : undefined,
      editColumn: editColumn && editColumn.trim() !== "" ? editColumn.trim() : undefined,
      editText: editText != null && String(editText).trim() !== "" ? editText : undefined
    },
    // Timestamp ensures the add-in can detect a new command even if
    // the same student is highlighted twice in a row
    timestamp: new Date().toISOString()
  };

  const commandJson = JSON.stringify(command);

  // Write to the custom document property
  workbook.getProperties().addCustomProperty("SRK_Command", commandJson);

  return `SUCCESS: Highlight command sent for student "${syStudentId}" (${new Date().toISOString()})`;
}
