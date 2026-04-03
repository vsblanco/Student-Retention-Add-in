/**
 * CreateLDASheets – Office Script for Power Automate
 *
 * Duplicates the "Master List" sheet and filters out rows that don't meet
 * the criteria. This preserves ALL formatting, conditional formatting,
 * colors, formulas, and column widths from the original — no rebuilding.
 *
 * Per-campus mode: creates one copy per campus, filters each to only
 * show that campus's students meeting the Days Out threshold.
 *
 * Single/date mode: creates one copy with all campuses filtered by Days Out.
 *
 * @param daysOut        - Days Out threshold (default: 5). Students with
 *                         Days Out >= this value are kept.
 * @param sheetNameMode  - "campus" for per-campus sheets, "date" for single
 *                         dated sheet (default: "campus")
 */
function main(
  workbook: ExcelScript.Workbook,
  daysOut?: number,
  sheetNameMode?: string
): string {
  const DAYS_OUT_THRESHOLD = daysOut ?? 5;
  const SHEET_MODE = sheetNameMode ?? "campus";

  // ── Helpers ──────────────────────────────────────────────────────────
  const stripStr = (s: string) => String(s || "").trim().toLowerCase().replace(/\s+/g, "");

  // ── 1. Validate Master List exists ───────────────────────────────────
  const masterSheet = workbook.getWorksheet("Master List");
  if (!masterSheet) return "ERROR: No 'Master List' sheet found.";

  const masterRange = masterSheet.getUsedRange();
  if (!masterRange) return "ERROR: Master List is empty.";

  const masterValues = masterRange.getValues() as (string | number | boolean)[][];
  if (masterValues.length < 2) return "ERROR: Master List has no data rows.";

  const headers = masterValues[0];

  // ── 2. Find key columns ──────────────────────────────────────────────
  function findCol(aliases: string[]): number {
    for (const alias of aliases) {
      const target = stripStr(alias);
      const idx = headers.findIndex((h) => stripStr(String(h)) === target);
      if (idx !== -1) return idx;
    }
    return -1;
  }

  const daysOutIdx = findCol(["Days Out"]);
  const campusIdx = findCol(["Campus"]);

  if (daysOutIdx === -1) return "ERROR: Could not find 'Days Out' column in Master List.";

  // ── 3. Determine campuses ────────────────────────────────────────────
  let campusList: string[] = [];

  if (SHEET_MODE === "campus" && campusIdx !== -1) {
    const campusSet = new Set<string>();
    for (let i = 1; i < masterValues.length; i++) {
      const val = String(masterValues[i][campusIdx] || "").trim();
      if (val) campusSet.add(val);
    }
    campusList = Array.from(campusSet).sort();
  }

  // Fallback: single sheet
  if (campusList.length === 0) {
    if (SHEET_MODE === "campus" && campusIdx !== -1) {
      for (let i = 1; i < masterValues.length; i++) {
        const val = String(masterValues[i][campusIdx] || "").trim();
        if (val) { campusList = [val]; break; }
      }
    }
    if (campusList.length === 0) {
      const today = new Date();
      campusList = [`LDA ${today.getMonth() + 1}-${today.getDate()}-${today.getFullYear()}`];
    }
  }

  // ── 4. Create one sheet per campus ───────────────────────────────────
  const existingSheets = workbook.getWorksheets().map((s) => s.getName());
  let sheetsCreated = 0;
  let totalStudents = 0;

  for (let ci = 0; ci < campusList.length; ci++) {
    const campusName = campusList[ci];

    // Generate unique sheet name
    let sheetName = campusName;
    let counter = 2;
    while (existingSheets.includes(sheetName)) {
      sheetName = `${campusName} (${counter++})`;
    }
    existingSheets.push(sheetName);

    // Duplicate the Master List
    const newSheet = masterSheet.copy(ExcelScript.WorksheetPositionType.after, masterSheet);
    newSheet.setName(sheetName);
    if (ci === 0) newSheet.activate();

    // Read the copied sheet's data
    const usedRange = newSheet.getUsedRange()!;
    const values = usedRange.getValues() as (string | number | boolean)[][];

    // ── Delete rows that don't match criteria (bottom-up to preserve indices)
    const rowsToDelete: number[] = [];

    for (let r = 1; r < values.length; r++) {
      const daysOutVal = values[r][daysOutIdx];
      const campusVal = campusIdx !== -1 ? String(values[r][campusIdx] || "").trim() : "";

      let shouldKeep = true;

      // Filter by Days Out: keep students >= threshold
      if (typeof daysOutVal !== "number" || daysOutVal < DAYS_OUT_THRESHOLD) {
        shouldKeep = false;
      }

      // Filter by campus (if multi-campus mode)
      if (shouldKeep && campusIdx !== -1 && campusList.length > 1) {
        if (campusVal !== campusName) {
          shouldKeep = false;
        }
      }

      if (!shouldKeep) {
        rowsToDelete.push(r);
      }
    }

    // Delete rows bottom-up to preserve row indices
    for (let i = rowsToDelete.length - 1; i >= 0; i--) {
      const rowIdx = rowsToDelete[i];
      const rowRange = newSheet.getRangeByIndexes(rowIdx, 0, 1, headers.length);
      rowRange.delete(ExcelScript.DeleteShiftDirection.up);
    }

    // Count remaining students (total rows minus header)
    const finalRange = newSheet.getUsedRange();
    const finalRowCount = finalRange ? finalRange.getRowCount() - 1 : 0;
    totalStudents += finalRowCount;

    // Auto-fit columns
    if (finalRange) {
      finalRange.getFormat().autofitColumns();
    }

    sheetsCreated++;
  }

  return `SUCCESS: Created ${sheetsCreated} LDA sheet(s) with ${totalStudents} total students across ${campusList.join(", ")}.`;
}
