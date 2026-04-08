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

  // ── Tab colors — one per campus, cycles through the palette ──────────
  const TAB_COLORS = [
    "#4472C4", // Blue
    "#ED7D31", // Orange
    "#70AD47", // Green
    "#FFC000", // Gold
    "#5B9BD5", // Light Blue
    "#A5A5A5", // Gray
    "#264478", // Dark Blue
    "#9B57A0", // Purple
    "#43682B", // Dark Green
    "#BF8F00", // Dark Gold
  ];

  // ── Default columns to SHOW (everything else gets hidden) ────────────
  const DEFAULT_VISIBLE_COLUMNS = [
    "assigned", "studentname", "student name", "gradebook", "grade book",
    "programversion", "program", "shift", "lda", "daysout", "days out",
    "grade", "missingassignments", "missing assignments", "outreach",
    "phone", "otherphone", "other phone", "campus",
  ];

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
  // We already have masterValues from step 1 — reuse it to decide which
  // rows to delete in each campus copy (no re-reading the copied sheet).
  const existingSheets = workbook.getWorksheets().map((s) => s.getName());
  let sheetsCreated = 0;
  let totalStudents = 0;
  const headerColCount = headers.length;

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

    // ── Determine which rows to delete using the cached masterValues ──
    const rowsToDelete: number[] = [];
    for (let r = 1; r < masterValues.length; r++) {
      const daysOutVal = masterValues[r][daysOutIdx];
      const campusVal = campusIdx !== -1 ? String(masterValues[r][campusIdx] || "").trim() : "";

      let shouldKeep = true;
      if (typeof daysOutVal !== "number" || daysOutVal < DAYS_OUT_THRESHOLD) {
        shouldKeep = false;
      }
      if (shouldKeep && campusIdx !== -1 && campusList.length > 1) {
        if (campusVal !== campusName) shouldKeep = false;
      }
      if (!shouldKeep) rowsToDelete.push(r);
    }

    // ── Batch delete: group consecutive rows into single range deletes ──
    // Process bottom-up, merging runs of consecutive rows into one range.
    const studentsKept = (masterValues.length - 1) - rowsToDelete.length;
    totalStudents += studentsKept;

    if (rowsToDelete.length > 0) {
      // Build runs of consecutive rows (walking from highest to lowest)
      const runs: { start: number; count: number }[] = [];
      let runEnd = rowsToDelete[rowsToDelete.length - 1];
      let runStart = runEnd;
      for (let i = rowsToDelete.length - 2; i >= 0; i--) {
        const r = rowsToDelete[i];
        if (r === runStart - 1) {
          runStart = r; // extend current run downward
        } else {
          runs.push({ start: runStart, count: runEnd - runStart + 1 });
          runEnd = r;
          runStart = r;
        }
      }
      runs.push({ start: runStart, count: runEnd - runStart + 1 });

      // Delete runs (already ordered bottom-up, so indices stay valid)
      for (const run of runs) {
        const rowRange = newSheet.getRangeByIndexes(run.start, 0, run.count, headerColCount);
        rowRange.delete(ExcelScript.DeleteShiftDirection.up);
      }
    }

    // Set tab color (cycles through palette)
    const tabColor = TAB_COLORS[ci % TAB_COLORS.length];
    newSheet.setTabColor(tabColor);

    // Auto-fit columns BEFORE hiding, so hidden state is preserved
    const finalRange = newSheet.getUsedRange();
    if (finalRange) {
      finalRange.getFormat().autofitColumns();
    }

    // Conditional format on Campus column — match tab color when cell equals campus name
    if (campusIdx !== -1 && studentsKept > 0) {
      const campusColRange = newSheet.getRangeByIndexes(1, campusIdx, studentsKept, 1);
      const cf = campusColRange.addConditionalFormat(ExcelScript.ConditionalFormatType.cellValue);
      cf.getCellValue().getFormat().getFill().setColor(tabColor);
      cf.getCellValue().setRule({
        formula1: `"${campusName.replace(/"/g, '""')}"`,
        operator: ExcelScript.ConditionalCellValueOperator.equalTo,
      });
    }

    // ── FINAL STEP: Hide columns not in the visible list ──
    // This must be the last step so nothing re-reveals the hidden columns.
    // Batch consecutive hidden columns into single range calls to reduce
    // API calls and avoid payload issues.
    const visibleSet = new Set(DEFAULT_VISIBLE_COLUMNS);
    const hiddenRuns: { start: number; count: number }[] = [];
    let runStart = -1;
    for (let c = 0; c < headerColCount; c++) {
      const headerStripped = stripStr(String(headers[c]));
      const shouldHide = !visibleSet.has(headerStripped);
      if (shouldHide) {
        if (runStart === -1) runStart = c;
      } else if (runStart !== -1) {
        hiddenRuns.push({ start: runStart, count: c - runStart });
        runStart = -1;
      }
    }
    if (runStart !== -1) {
      hiddenRuns.push({ start: runStart, count: headerColCount - runStart });
    }
    for (const run of hiddenRuns) {
      newSheet.getRangeByIndexes(0, run.start, 1, run.count).getEntireColumn().setColumnHidden(true);
    }

    sheetsCreated++;
  }

  return `SUCCESS: Created ${sheetsCreated} LDA sheet(s) with ${totalStudents} total students across ${campusList.join(", ")}.`;
}
