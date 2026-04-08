/**
 * importMasterList – Office Script for Power Automate
 *
 * Replicates the Student Retention Add-in's master list import logic.
 * Accepts a JSON payload (headers + data rows) and writes it to a "Master List"
 * sheet, preserving Assigned colors, Gradebook hyperlinks, and unmatched columns
 * from any prior import. New students are highlighted in light blue.
 *
 * Power Automate usage:
 *   Pass the workbook and a JSON string via the "Run script" action.
 *   The JSON shape must be: { headers: string[], data: (string|number)[][] }
 */
function main(
  workbook: ExcelScript.Workbook,
  jsonPayload: string
): string {
  // ── 1. Parse & validate payload ──────────────────────────────────────
  interface Payload {
    headers: string[];
    data: (string | number | null)[][];
  }

  let payload: Payload;
  try {
    payload = JSON.parse(jsonPayload) as Payload;
  } catch {
    return "ERROR: Could not parse JSON payload.";
  }

  if (!payload.headers || !payload.data || payload.data.length === 0) {
    return "ERROR: Payload must contain 'headers' (string[]) and 'data' (any[][]).";
  }

  const incomingHeaders = payload.headers;
  const incomingData = payload.data;

  // ── 2. Constants – column alias maps (mirrors shared-utilities.js) ───
  const STUDENT_NAME_COLS = ["student name", "studentname", "student"];
  const STUDENT_ID_COLS = ["student id", "systudentid", "id"];
  const GRADEBOOK_COLS = [
    "grade book", "gradebook", "gradebooklink", "gradelink"
  ];
  const ASSIGNED_COLS = ["assigned", "advisor"];
  const MISSING_ASSIGNMENTS_COLS = [
    "missing assignments", "missingassignments",
    "course missing assignments", "total missing"
  ];
  const GRADE_COLS = ["grade", "current score", "course grade", "grades"];
  const LAST_COURSE_GRADE_COLS = ["last course grade", "lastcoursegrade"];
  const HOLD_COLS = ["hold"];
  const ATTENDANCE_COLS = ["attendance %", "attendance%", "attendancepercent", "attendance"];
  const EXPECTED_START_COLS = ["expected start date", "start date", "expstartdate"];
  const PROGRAM_VERSION_COLS = ["programversion", "program version", "program", "progversdescrip"];

  // ── helpers ──────────────────────────────────────────────────────────
  const normalize = (h: string) => h.toLowerCase().replace(/\s+/g, "");

  function findColIndex(headers: string[], aliases: string[]): number {
    const lower = headers.map((h) => h.toLowerCase());
    for (const alias of aliases) {
      const idx = lower.indexOf(alias);
      if (idx !== -1) return idx;
    }
    // fallback: normalized match
    const norm = headers.map(normalize);
    for (const alias of aliases) {
      const idx = norm.indexOf(normalize(alias));
      if (idx !== -1) return idx;
    }
    return -1;
  }

  function normalizeName(name: string): string {
    if (!name) return "";
    let n = name.trim().toLowerCase();
    if (n.includes(",")) {
      const parts = n.split(",").map((p) => p.trim());
      if (parts.length > 1) return `${parts[1]} ${parts[0]}`;
    }
    return n;
  }

  function formatToLastFirst(name: string): string {
    if (!name) return "";
    let n = name.trim();
    if (n.includes(",")) {
      return n
        .split(",")
        .map((p) => p.trim())
        .join(", ");
    }
    const parts = n.split(" ").filter((p) => p);
    if (parts.length > 1) {
      const last = parts.pop()!;
      return `${last}, ${parts.join(" ")}`;
    }
    return n;
  }

  function parseDate(v: string | number | null): Date | null {
    if (v == null) return null;
    if (typeof v === "number" && v > 25569) {
      return new Date((v - 25569) * 86400 * 1000);
    }
    if (typeof v === "string") {
      const d = new Date(v);
      if (!isNaN(d.getTime())) return d;
    }
    return null;
  }

  // ── 3. Get or create the Master List sheet ───────────────────────────
  let sheet = workbook.getWorksheet("Master List");
  const isNewSheet = !sheet;

  if (!sheet) {
    sheet = workbook.addWorksheet("Master List");
    // Write header row for brand-new sheet
    const headerRange = sheet.getRangeByIndexes(0, 0, 1, incomingHeaders.length);
    headerRange.setValues([incomingHeaders]);
    headerRange.getFormat().getFont().setBold(true);
  }

  // ── 4. Read existing sheet data ──────────────────────────────────────
  const usedRange = sheet.getUsedRange();
  let existingValues: (string | number | boolean)[][] = [];
  let existingFormulas: string[][] = [];
  let masterHeaders: string[] = [];

  if (usedRange) {
    existingValues = usedRange.getValues() as (string | number | boolean)[][];
    existingFormulas = usedRange.getFormulas() as string[][];
    masterHeaders = existingValues[0].map((h) => String(h ?? ""));
  } else {
    masterHeaders = [...incomingHeaders];
  }

  const lowerMaster = masterHeaders.map((h) => h.toLowerCase());
  const normMaster = masterHeaders.map(normalize);
  const lowerIncoming = incomingHeaders.map((h) => h.toLowerCase());
  const normIncoming = incomingHeaders.map(normalize);

  // ── 5. Column mapping (incoming → master) ────────────────────────────
  const colMapping: number[] = lowerIncoming.map((inc, idx) => {
    let mIdx = lowerMaster.indexOf(inc);
    if (mIdx === -1) mIdx = normMaster.indexOf(normIncoming[idx]);
    return mIdx;
  });

  // Detect & append new columns
  const newColumns: string[] = [];
  for (let i = 0; i < incomingHeaders.length; i++) {
    if (colMapping[i] === -1) {
      const newIdx = masterHeaders.length;
      masterHeaders.push(incomingHeaders[i]);
      lowerMaster.push(lowerIncoming[i]);
      normMaster.push(normIncoming[i]);
      colMapping[i] = newIdx;
      newColumns.push(incomingHeaders[i]);
    }
  }

  // ── 6. Identify key columns ──────────────────────────────────────────
  const inNameCol = findColIndex(incomingHeaders, STUDENT_NAME_COLS);
  const mNameCol = findColIndex(masterHeaders, STUDENT_NAME_COLS);
  const mGradebookCol = findColIndex(masterHeaders, GRADEBOOK_COLS);
  const mAssignedCol = findColIndex(masterHeaders, ASSIGNED_COLS);
  const mMissingCol = findColIndex(masterHeaders, MISSING_ASSIGNMENTS_COLS);
  const inMissingCol = findColIndex(incomingHeaders, MISSING_ASSIGNMENTS_COLS);
  const mIdCol = findColIndex(masterHeaders, STUDENT_ID_COLS);
  const inIdCol = findColIndex(incomingHeaders, STUDENT_ID_COLS);
  const mExpStartCol = findColIndex(masterHeaders, EXPECTED_START_COLS);
  const mGradeCol = findColIndex(masterHeaders, GRADE_COLS);
  const mLastGradeCol = findColIndex(masterHeaders, LAST_COURSE_GRADE_COLS);
  const mHoldCol = findColIndex(masterHeaders, HOLD_COLS);
  const mAttendanceCol = findColIndex(masterHeaders, ATTENDANCE_COLS);
  const mProgramVersionCol = findColIndex(masterHeaders, PROGRAM_VERSION_COLS);

  if (inNameCol === -1) return "ERROR: Incoming data has no Student Name column.";
  if (mNameCol === -1) return "ERROR: Master List has no Student Name column.";

  // ── 7. Build preservation maps from existing data ────────────────────
  const masterDataMap = new Map<
    string,
    { gradebookFormula: string | null; assigned: string | null }
  >();
  const valueToColorMap = new Map<string, string>();

  // Determine unmatched master columns (those not covered by incoming data)
  const matchedSet = new Set(colMapping);
  const unmatchedCols: number[] = [];
  for (let i = 0; i < masterHeaders.length; i++) {
    if (!matchedSet.has(i)) unmatchedCols.push(i);
  }

  // Non-blank unmatched columns
  const nonBlankUnmatchedCols: number[] = [];
  if (!isNewSheet && existingValues.length > 1) {
    for (const ci of unmatchedCols) {
      for (let r = 1; r < existingValues.length; r++) {
        if (
          ci < existingValues[r].length &&
          existingValues[r][ci] != null &&
          String(existingValues[r][ci]).trim() !== ""
        ) {
          nonBlankUnmatchedCols.push(ci);
          break;
        }
      }
    }
  }

  // Preservation map (for unmatched columns with data)
  const preservationMap = new Map<
    string,
    { values: Record<number, string | number | boolean>; formulas: Record<number, string> }
  >();

  if (!isNewSheet && existingValues.length > 1) {
    for (let r = 1; r < existingValues.length; r++) {
      const name = existingValues[r][mNameCol];
      if (!name) continue;
      const nName = normalizeName(String(name));

      // Gradebook + Assigned
      const gbFormula =
        mGradebookCol !== -1 && existingFormulas[r][mGradebookCol]
          ? String(existingFormulas[r][mGradebookCol])
          : null;
      const assignedVal =
        mAssignedCol !== -1 ? String(existingValues[r][mAssignedCol] ?? "") : null;

      masterDataMap.set(nName, {
        gradebookFormula:
          gbFormula && gbFormula.startsWith("=") ? gbFormula : null,
        assigned: assignedVal && assignedVal.trim() !== "" ? assignedVal : null,
      });

      // Unmatched column preservation (keyed by ID then name)
      if (nonBlankUnmatchedCols.length > 0) {
        let key: string | null = null;
        if (mIdCol !== -1) {
          const idVal = existingValues[r][mIdCol];
          if (idVal != null && String(idVal).trim() !== "") key = String(idVal).trim();
        }
        if (!key) key = nName;

        const pValues: Record<number, string | number | boolean> = {};
        const pFormulas: Record<number, string> = {};
        for (const ci of nonBlankUnmatchedCols) {
          if (ci >= existingValues[r].length) continue;
          const formula = String(existingFormulas[r][ci] ?? "");
          const val = existingValues[r][ci];
          if (formula.startsWith("=")) {
            pFormulas[ci] = formula;
            pValues[ci] = val;
          } else if (val != null && String(val).trim() !== "") {
            pValues[ci] = val;
          }
        }
        preservationMap.set(key, { values: pValues, formulas: pFormulas });
      }
    }

    // Sample Assigned column colors (one cell per unique value)
    if (mAssignedCol !== -1) {
      const seen = new Set<string>();
      for (let r = 1; r < existingValues.length; r++) {
        const val = String(existingValues[r][mAssignedCol] ?? "").trim();
        if (!val || seen.has(val)) continue;
        seen.add(val);
        const cell = sheet!.getCell(r, mAssignedCol);
        const color = cell.getFormat().getFill().getColor();
        if (color && color !== "#FFFFFF" && color !== "#000000") {
          valueToColorMap.set(val, color);
        }
      }
    }
  }

  // ── 8. Categorize students ───────────────────────────────────────────
  const newStudents: (string | number | null)[][] = [];
  const existingStudents: (string | number | null)[][] = [];

  for (const row of incomingData) {
    const nName = normalizeName(String(row[inNameCol] ?? ""));
    if (masterDataMap.has(nName)) {
      existingStudents.push(row);
    } else {
      newStudents.push(row);
    }
  }

  // ── 9. Clear existing data (keep header row) ────────────────────────
  if (!isNewSheet && existingValues.length > 1) {
    const clearRange = sheet!.getRangeByIndexes(
      1, 0,
      existingValues.length - 1,
      Math.max(masterHeaders.length, existingValues[0].length)
    );
    clearRange.clear(ExcelScript.ClearApplyTo.all);
  }

  // Update header row (may have new columns)
  const headerRange = sheet!.getRangeByIndexes(0, 0, 1, masterHeaders.length);
  headerRange.setValues([masterHeaders]);
  headerRange.getFormat().getFont().setBold(true);

  // ── 10. Build output rows ────────────────────────────────────────────
  const allStudents = [...newStudents, ...existingStudents];
  if (allStudents.length === 0) return "No students to import.";

  const outputValues: (string | number | boolean)[][] = [];
  const outputFormulas: (string | null)[][] = [];
  const cellColors: { row: number; col: number; color: string }[] = [];

  for (let i = 0; i < allStudents.length; i++) {
    const inRow = allStudents[i];
    const outRow: (string | number | boolean)[] = new Array(masterHeaders.length).fill("");
    const fRow: (string | null)[] = new Array(masterHeaders.length).fill(null);

    // Map incoming → master
    for (let c = 0; c < inRow.length; c++) {
      const mc = colMapping[c];
      if (mc === -1) continue;
      let val = inRow[c] ?? "";

      // Format name to "Last, First"
      if (mc === mNameCol) val = formatToLastFirst(String(val));

      // Trim Program Version: remove prefix up to and including the first 4-digit year
      // e.g., "B2-Q-2024 Allied Health Management" → "Allied Health Management"
      if (mc === mProgramVersionCol && val) {
        const match = String(val).match(/^.*?\d{4}\s+(.*)$/);
        if (match) val = match[1].trim();
      }

      // Wrap Gradebook URLs in HYPERLINK
      if (mc === mGradebookCol && val) {
        const url = String(val).trim();
        if (url.startsWith("http://") || url.startsWith("https://")) {
          fRow[mc] = `=HYPERLINK("${url}", "Grade Book")`;
          val = "Grade Book";
        }
      }

      outRow[mc] = val;
    }

    // Missing Assignments: default to 0 when gradebook link exists
    if (mMissingCol !== -1) {
      const hasGB = mGradebookCol !== -1 && outRow[mGradebookCol];
      const inVal = inMissingCol !== -1 ? inRow[inMissingCol] : null;
      outRow[mMissingCol] = hasGB
        ? inVal != null && inVal !== "" ? inVal : 0
        : "";
    }

    // Preserve existing Gradebook & Assigned for returning students
    const nName = normalizeName(String(inRow[inNameCol] ?? ""));
    if (masterDataMap.has(nName)) {
      const existing = masterDataMap.get(nName)!;
      if (existing.gradebookFormula && mGradebookCol !== -1 && !outRow[mGradebookCol]) {
        fRow[mGradebookCol] = existing.gradebookFormula;
        const match = existing.gradebookFormula.match(/, *"([^"]+)"\)/i);
        outRow[mGradebookCol] = match ? match[1] : "Gradebook";
      }
      if (existing.assigned && mAssignedCol !== -1 && !outRow[mAssignedCol]) {
        outRow[mAssignedCol] = existing.assigned;
      }
    }

    // Restore unmatched-column data
    if (nonBlankUnmatchedCols.length > 0) {
      let pKey: string | null = null;
      if (inIdCol !== -1) {
        const id = inRow[inIdCol];
        if (id != null && String(id).trim() !== "") pKey = String(id).trim();
      }
      if (!pKey) pKey = nName;
      if (pKey && preservationMap.has(pKey)) {
        const p = preservationMap.get(pKey)!;
        for (const ci of nonBlankUnmatchedCols) {
          if (p.formulas[ci]) {
            fRow[ci] = p.formulas[ci];
            outRow[ci] = p.values[ci] ?? "";
          } else if (p.values[ci] !== undefined) {
            outRow[ci] = p.values[ci];
          }
        }
      }
    }

    // Queue Assigned cell color
    if (mAssignedCol !== -1) {
      const aVal = String(outRow[mAssignedCol] ?? "").trim();
      if (aVal && valueToColorMap.has(aVal)) {
        cellColors.push({ row: i + 1, col: mAssignedCol, color: valueToColorMap.get(aVal)! });
      }
    }

    outputValues.push(outRow);
    outputFormulas.push(fRow);
  }

  // ── 11. Write data to sheet ──────────────────────────────────────────
  const dataRange = sheet!.getRangeByIndexes(1, 0, outputValues.length, masterHeaders.length);
  dataRange.setValues(outputValues);

  // Apply formulas where they exist
  for (let r = 0; r < outputFormulas.length; r++) {
    for (let c = 0; c < outputFormulas[r].length; c++) {
      if (outputFormulas[r][c]) {
        sheet!.getCell(r + 1, c).setFormula(outputFormulas[r][c]!);
      }
    }
  }

  // ── 12. Highlight preserved columns in light gray ────────────────────
  if (nonBlankUnmatchedCols.length > 0) {
    for (const ci of nonBlankUnmatchedCols) {
      const colRange = sheet!.getRangeByIndexes(1, ci, outputValues.length, 1);
      colRange.getFormat().getFill().setColor("#EDEDED");
    }
  }

  // ── 13. Highlight new students in light blue ─────────────────────────
  if (mExpStartCol !== -1) {
    // Highlight students with the latest Expected Start Date
    let latestDate: Date | null = null;
    for (const row of outputValues) {
      const d = parseDate(row[mExpStartCol] as string | number | null);
      if (d && (!latestDate || d > latestDate)) latestDate = d;
    }
    if (latestDate) {
      const latestStr = latestDate.toDateString();
      for (let i = 0; i < outputValues.length; i++) {
        const d = parseDate(outputValues[i][mExpStartCol] as string | number | null);
        if (d && d.toDateString() === latestStr) {
          sheet!
            .getRangeByIndexes(i + 1, 0, 1, masterHeaders.length)
            .getFormat()
            .getFill()
            .setColor("#ADD8E6");
        }
      }
    }
  } else if (newStudents.length > 0) {
    const hlRange = sheet!.getRangeByIndexes(1, 0, newStudents.length, masterHeaders.length);
    hlRange.getFormat().getFill().setColor("#ADD8E6");
  }

  // ── 14. Apply preserved Assigned cell colors ─────────────────────────
  for (const cc of cellColors) {
    sheet!.getCell(cc.row, cc.col).getFormat().getFill().setColor(cc.color);
  }

  // ── 15. Conditional formatting – clear old rules, then apply ─────────
  // Clear all existing conditional formats on the sheet at once
  const cfRange = sheet!.getUsedRange();
  if (cfRange) {
    cfRange.clearAllConditionalFormats();
  }

  // Grade column: Red → Yellow → Green color scale
  applyThreeColorScale(sheet!, masterHeaders, mGradeCol, outputValues);

  // Last Course Grade: same color scale
  applyThreeColorScale(sheet!, masterHeaders, mLastGradeCol, outputValues);

  // Missing Assignments: 0 → light green
  if (mMissingCol !== -1 && outputValues.length > 0) {
    const maRange = sheet!.getRangeByIndexes(1, mMissingCol, outputValues.length, 1);
    const cfMissing = maRange.addConditionalFormat(ExcelScript.ConditionalFormatType.cellValue);
    cfMissing.getCellValue().getFormat().getFill().setColor("#E2EFDA");
    cfMissing.getCellValue().setRule({
      formula1: "0",
      operator: ExcelScript.ConditionalCellValueOperator.equalTo,
    });
  }

  // Hold: "Yes" → light red
  if (mHoldCol !== -1 && outputValues.length > 0) {
    const holdRange = sheet!.getRangeByIndexes(1, mHoldCol, outputValues.length, 1);
    const cfHold = holdRange.addConditionalFormat(ExcelScript.ConditionalFormatType.containsText);
    cfHold.getTextComparison().getFormat().getFill().setColor("#FFB6C1");
    cfHold.getTextComparison().setRule({
      operator: ExcelScript.ConditionalTextOperator.contains,
      text: "Yes",
    });
  }

  // Attendance: 3-color scale (red-yellow-green)
  if (mAttendanceCol !== -1 && outputValues.length > 0) {
    applyThreeColorScale(sheet!, masterHeaders, mAttendanceCol, outputValues);
  }

  // ── 16. Auto-fit columns ─────────────────────────────────────────────
  sheet!.getUsedRange()?.getFormat().autofitColumns();

  return `SUCCESS: Imported ${allStudents.length} students (${newStudents.length} new, ${existingStudents.length} existing).${newColumns.length > 0 ? ` Added ${newColumns.length} new column(s): ${newColumns.join(", ")}` : ""}`;
}

/**
 * Applies a Red → Yellow → Green 3-color scale to a numeric column.
 * Auto-detects 0-1 vs 0-100 scale.
 */
function applyThreeColorScale(
  sheet: ExcelScript.Worksheet,
  headers: string[],
  colIdx: number,
  data: (string | number | boolean)[][]
): void {
  if (colIdx === -1 || data.length === 0) return;

  const range = sheet.getRangeByIndexes(1, colIdx, data.length, 1);

  // Detect scale
  let isPercent = false;
  for (let i = 0; i < Math.min(data.length, 10); i++) {
    const v = data[i][colIdx];
    if (typeof v === "number" && v > 1) {
      isPercent = true;
      break;
    }
  }

  const cf = range.addConditionalFormat(ExcelScript.ConditionalFormatType.colorScale);
  cf.getColorScale().setCriteria({
    minimum: {
      type: ExcelScript.ConditionalFormatColorCriterionType.lowestValue,
      color: "#F8696B",
    },
    midpoint: {
      type: ExcelScript.ConditionalFormatColorCriterionType.number,
      formula: isPercent ? "70" : "0.7",
      color: "#FFEB84",
    },
    maximum: {
      type: ExcelScript.ConditionalFormatColorCriterionType.highestValue,
      color: "#63BE7B",
    },
  });
}
