/**
 * HighlightRows – Office Script for Power Automate
 *
 * Two-path approach for instant highlighting in shared workbooks:
 *
 *   1. DIRECT: Applies formatting to the file (persists even if nobody
 *      has the workbook open — serves as the durable fallback)
 *   2. REAL-TIME: Writes the command to the "SRK_Command" custom document
 *      property, which syncs via co-authoring within seconds. The Office
 *      Add-in (if active) polls this property and executes the highlight
 *      locally in the user's session for instant visibility.
 *
 * Only one add-in instance will process the command — the first to claim
 * it writes its session ID to "SRK_CommandClaim", and others skip.
 *
 * @param syStudentId - The student's ID to find in the sheet
 * @param targetSheet - Name of the worksheet to highlight in
 * @param startCol    - Start column (name or 0-based index) for the highlight range
 * @param endCol      - End column (name or 0-based index) for the highlight range
 * @param color       - Hex color for the highlight (optional, omit to remove)
 * @param editColumn  - Optional column (name or index) to write text into
 * @param editText    - Optional text to write into editColumn
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
  // ── Defaults ─────────────────────────────────────────────────────────
  // No color provided = "remove mode"; with color = "apply mode"
  const isRemoveMode = !color || color.trim() === "";
  const highlightColor = isRemoveMode ? "" : color!.trim();

  // ── Validate required params ─────────────────────────────────────────
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

  // ── ID column aliases (same as the add-in) ──────────────────────────
  const ID_ALIASES = ["student id", "systudentid", "student identifier", "id"];

  // ── 1. Find the target sheet ─────────────────────────────────────────
  let sheet = workbook.getWorksheet(targetSheet.trim());

  // If exact match fails, try date-format normalization (01/05/2026 vs 1/5/2026)
  if (!sheet) {
    const variations = normalizeDateFormat(targetSheet.trim());
    const allSheets = workbook.getWorksheets();
    for (const variant of variations) {
      const match = allSheets.find((s) => s.getName() === variant);
      if (match) {
        sheet = match;
        break;
      }
    }
  }

  if (!sheet) {
    return `ERROR: Sheet "${targetSheet}" not found in the workbook.`;
  }

  // ── 2. Read the sheet ────────────────────────────────────────────────
  const usedRange = sheet.getUsedRange();
  if (!usedRange) {
    return `ERROR: Sheet "${targetSheet}" is empty.`;
  }

  const values = usedRange.getValues() as (string | number | boolean)[][];
  if (values.length < 2) {
    return `ERROR: Sheet "${targetSheet}" has no data rows.`;
  }

  const headers = values[0].map((h) => String(h ?? "").trim());
  const headersLower = headers.map((h) => h.toLowerCase());

  // ── 3. Resolve column references (name → index) ─────────────────────
  function resolveCol(col: string, paramName: string): { index: number; error: string | null } {
    // Try as a number first
    const num = Number(col);
    if (!isNaN(num) && String(num) === col.trim()) {
      if (num < 0 || num >= headers.length) {
        return { index: -1, error: `${paramName} index (${num}) is out of range (0-${headers.length - 1}).` };
      }
      return { index: num, error: null };
    }
    // Find by name (case-insensitive)
    const colLower = col.trim().toLowerCase();
    const idx = headersLower.indexOf(colLower);
    if (idx === -1) {
      return { index: -1, error: `${paramName} column "${col}" not found in sheet headers.` };
    }
    return { index: idx, error: null };
  }

  const startResult = resolveCol(startCol, "startCol");
  if (startResult.error) return `ERROR: ${startResult.error}`;

  const endResult = resolveCol(endCol, "endCol");
  if (endResult.error) return `ERROR: ${endResult.error}`;

  const startColIndex = Math.min(startResult.index, endResult.index);
  const endColIndex = Math.max(startResult.index, endResult.index);

  let editColIndex = -1;
  if (editColumn != null && editColumn.trim() !== "") {
    const editResult = resolveCol(editColumn, "editColumn");
    if (editResult.error) return `ERROR: ${editResult.error}`;
    editColIndex = editResult.index;
  }

  // ── 4. Find the Student ID column ────────────────────────────────────
  let idColIndex = -1;
  for (const alias of ID_ALIASES) {
    const idx = headersLower.indexOf(alias);
    if (idx !== -1) {
      idColIndex = idx;
      break;
    }
  }
  if (idColIndex === -1) {
    return `ERROR: No Student ID column found in sheet "${targetSheet}". Expected one of: ${ID_ALIASES.join(", ")}.`;
  }

  // ── 5. Find the student's row ────────────────────────────────────────
  const targetId = syStudentId.trim();
  let targetRowIndex = -1;

  for (let r = 1; r < values.length; r++) {
    if (String(values[r][idColIndex]).trim() === targetId) {
      targetRowIndex = r;
      break;
    }
  }

  if (targetRowIndex === -1) {
    return `ERROR: Student ID "${syStudentId}" not found in sheet "${targetSheet}".`;
  }

  // ── 6. Apply or remove the highlight ──────────────────────────────────
  const columnCount = endColIndex - startColIndex + 1;
  const highlightRange = sheet.getRangeByIndexes(targetRowIndex, startColIndex, 1, columnCount);
  const currentColor = highlightRange.getFormat().getFill().getColor();
  const hasExistingHighlight = currentColor && currentColor !== "#FFFFFF" && currentColor !== "";

  if (isRemoveMode) {
    // No color provided → remove existing highlight (skip if nothing to remove)
    if (!hasExistingHighlight) {
      return `SKIPPED: Student "${targetId}" on "${sheet.getName()}" has no highlight to remove.`;
    }
    highlightRange.getFormat().getFill().clear();
  } else {
    // Color provided → apply highlight (skip if already the same color)
    if (hasExistingHighlight && currentColor!.toLowerCase() === highlightColor.toLowerCase()) {
      return `SKIPPED: Student "${targetId}" on "${sheet.getName()}" already highlighted with ${highlightColor}.`;
    }
    highlightRange.getFormat().getFill().setColor(highlightColor);
  }

  // ── 7. Optionally write text to editColumn ───────────────────────────
  let editMsg = "";
  if (editColIndex !== -1 && editText != null) {
    sheet.getCell(targetRowIndex, editColIndex).setValue(editText);
    editMsg = ` Wrote "${editText}" to column "${editColumn}".`;
  }

  // ── 8. Send real-time command via custom document property ─────────────
  // This syncs to active add-in sessions within seconds via co-authoring
  const command: Record<string, unknown> = {
    type: "SRK_HIGHLIGHT_STUDENT_ROW",
    data: {
      syStudentId: targetId,
      targetSheet: sheet.getName(),
      startCol: startCol.trim(),
      endCol: endCol.trim(),
      color: isRemoveMode ? undefined : highlightColor,
      editColumn: editColumn && editColumn.trim() !== "" ? editColumn.trim() : undefined,
      editText: editText != null && String(editText).trim() !== "" ? editText : undefined
    },
    timestamp: new Date().toISOString()
  };
  workbook.getProperties().addCustomProperty("SRK_Command", JSON.stringify(command));

  const action = isRemoveMode ? "Removed highlight from" : `Highlighted`;
  return `SUCCESS: ${action} student "${targetId}" on "${sheet.getName()}" (row ${targetRowIndex + 1}, columns ${startColIndex}-${endColIndex}).${editMsg}`;
}

/**
 * Generates date format variations to handle sheet name mismatches
 * like "01/05/2026" vs "1/5/2026" vs "01-05-2026".
 * Mirrors chromeExtensionService.js normalizeDateFormat().
 */
function normalizeDateFormat(dateStr: string): string[] {
  const datePattern = /^(.*?)(\d{1,2}[/-]\d{1,2}[/-]\d{4})$/;
  const match = dateStr.match(datePattern);
  if (!match) return [dateStr];

  const prefix = match[1];
  const datePart = match[2];
  const sep = datePart.includes("/") ? "/" : "-";
  const parts = datePart.split(sep);
  if (parts.length !== 3) return [dateStr];

  const [month, day, year] = parts;
  const variations = new Set<string>();
  variations.add(dateStr);

  for (const s of ["/", "-"]) {
    variations.add(`${prefix}${month.padStart(2, "0")}${s}${day.padStart(2, "0")}${s}${year}`);
    variations.add(`${prefix}${parseInt(month, 10)}${s}${parseInt(day, 10)}${s}${year}`);
    variations.add(`${prefix}${month.padStart(2, "0")}${s}${parseInt(day, 10)}${s}${year}`);
    variations.add(`${prefix}${parseInt(month, 10)}${s}${day.padStart(2, "0")}${s}${year}`);
  }

  return Array.from(variations);
}
