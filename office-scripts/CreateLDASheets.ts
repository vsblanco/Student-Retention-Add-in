/**
 * CreateLDASheets – Office Script for Power Automate
 *
 * Reads the "Master List" sheet and creates per-campus LDA sheets,
 * replicating the add-in's LDA processor logic:
 *
 *   - Filters students by Days Out threshold
 *   - Optionally creates a Failing Students table (grade < 60, days out <= 4)
 *   - Optionally creates a Low Attendance table (attendance < 60%)
 *   - Splits by campus (one sheet per campus) or creates a single dated sheet
 *   - Preserves value-based cell colors from Master List (advisor colors, etc.)
 *   - Copies conditional formatting from Master List columns
 *   - Creates Excel tables with TableStyleLight9
 *   - Hides implicit columns not in the configured column list
 *   - Auto-fits columns
 *
 * NOTE: DNC/LDA retention tags are NOT included in this version.
 *
 * @param daysOut            - Days Out threshold (default: 5)
 * @param includeFailingList - Include failing students table (default: false)
 * @param includeAttendanceList - Include low attendance table (default: false)
 * @param sheetNameMode      - "campus" for per-campus sheets, "date" for single dated sheet (default: "campus")
 * @param columnsJson        - JSON array of column config objects (optional, auto-detects if omitted)
 */
function main(
  workbook: ExcelScript.Workbook,
  daysOut?: number,
  includeFailingList?: boolean,
  includeAttendanceList?: boolean,
  sheetNameMode?: string,
  columnsJson?: string
): string {
  // ── Settings ─────────────────────────────────────────────────────────
  const DAYS_OUT_THRESHOLD = daysOut ?? 5;
  const INCLUDE_FAILING = includeFailingList ?? false;
  const INCLUDE_ATTENDANCE = includeAttendanceList ?? false;
  const SHEET_MODE = sheetNameMode ?? "campus";
  const BATCH_SIZE = 500;

  // ── Default column config (mirrors DefaultSettings.jsx) ──────────────
  interface ColConfig {
    name: string;
    alias?: string[];
    hidden?: boolean;
  }

  const DEFAULT_COLUMNS: ColConfig[] = [
    { name: "Assigned", alias: ["advisor"] },
    { name: "Student Name", alias: ["Student"] },
    { name: "Gradebook", alias: ["Gradelink", "gradeBookLink", "Grade Book"] },
    { name: "ProgramVersion", alias: ["Program", "ProgVersDescrip"] },
    { name: "Shift", alias: ["ShiftDescrip"] },
    { name: "LDA", alias: ["Last Date of Attendance", "Date of Attendance", "CurrentLDA"] },
    { name: "Days Out", alias: ["Days Out"] },
    { name: "Grade", alias: ["Course Grade", "current score", "current grade"] },
    { name: "Missing Assignments", alias: ["Total Missing"] },
    { name: "Outreach", alias: ["Comments", "Comment"] },
    { name: "Phone", alias: ["Phone Number", "Contact Number"] },
    { name: "Other Phone", alias: ["Second Phone", "Alt Phone"] },
  ];

  let configColumns: ColConfig[] = DEFAULT_COLUMNS;
  if (columnsJson && columnsJson.trim() !== "") {
    try {
      configColumns = JSON.parse(columnsJson) as ColConfig[];
    } catch {
      // Fall back to defaults
    }
  }

  // ── Helpers ──────────────────────────────────────────────────────────
  const stripStr = (s: string) => String(s || "").trim().toLowerCase().replace(/\s+/g, "");

  // ── 1. Read Master List ──────────────────────────────────────────────
  const masterSheet = workbook.getWorksheet("Master List");
  if (!masterSheet) return "ERROR: No 'Master List' sheet found.";

  const masterRange = masterSheet.getUsedRange();
  if (!masterRange) return "ERROR: Master List is empty.";

  const masterValues = masterRange.getValues() as (string | number | boolean)[][];
  const masterFormulas = masterRange.getFormulas() as string[][];
  if (masterValues.length < 2) return "ERROR: Master List has no data rows.";

  const headers = masterValues[0];

  // ── Column index resolver (space-insensitive, alias-aware) ───────────
  function getColIndex(settingName: string): number {
    const targetStripped = stripStr(settingName);

    // Find config for this column
    const colConfig = configColumns.find((c) => stripStr(c.name) === targetStripped);
    let aliases: string[] = [];
    if (colConfig && colConfig.alias) {
      aliases = colConfig.alias;
    }
    const candidates = [settingName, ...aliases];

    for (const rawCand of candidates) {
      const candStripped = stripStr(rawCand);
      const idx = headers.findIndex((h) => stripStr(String(h)) === candStripped);
      if (idx !== -1) return idx;
    }
    return -1;
  }

  // ── 2. Build output columns (configured + implicit hidden) ───────────
  const matchedColumns = configColumns.filter((col) => getColIndex(col.name) !== -1);
  const usedMasterIndices = new Set<number>();
  matchedColumns.forEach((col) => {
    const idx = getColIndex(col.name);
    if (idx !== -1) usedMasterIndices.add(idx);
  });

  const outputColumns: ColConfig[] = [...matchedColumns];
  headers.forEach((h, idx) => {
    if (!usedMasterIndices.has(idx) && h && String(h).trim() !== "") {
      outputColumns.push({ name: String(h), hidden: true });
    }
  });

  // ── 3. Key column indices ────────────────────────────────────────────
  const daysOutIdx = getColIndex("Days Out");
  const gradeIdx = getColIndex("Grade");
  const campusIdx = getColIndex("Campus");
  const outreachColIndex = outputColumns.findIndex((c) => stripStr(c.name) === "outreach");

  // Attendance column
  let attendanceIdx = getColIndex("Attendance %");
  if (attendanceIdx === -1) {
    attendanceIdx = headers.findIndex((h) => {
      const s = stripStr(String(h));
      return s === "attendance%" || s === "attendancepercent" || s === "attendance";
    });
  }

  // Missing Assignments column
  let missingIdx = getColIndex("Missing Assignments");
  if (missingIdx === -1) {
    missingIdx = headers.findIndex((h) => String(h).trim().toLowerCase().includes("missing"));
  }

  // Student ID column
  let studentIdIdx = getColIndex("Student Number");
  if (studentIdIdx === -1) {
    studentIdIdx = headers.findIndex((h) => {
      const s = stripStr(String(h));
      return s === "studentnumber" || s === "studentidentifier";
    });
  }

  if (daysOutIdx === -1) return "ERROR: Could not find 'Days Out' column in Master List.";

  // ── 4. Detect date columns (for number format) ──────────────────────
  const dateColumnNames = new Set<string>();
  outputColumns.forEach((colConfig) => {
    const masterIdx = getColIndex(colConfig.name);
    if (masterIdx === -1) return;
    const excelHeader = String(headers[masterIdx] || "").toLowerCase();
    if (/id|no\.|num|code|zip|postal|social|ssn|phone|grade|score|credit|fee|days|count/i.test(excelHeader)) return;
    let dateCount = 0;
    let numCount = 0;
    const limit = Math.min(masterValues.length, 100);
    for (let i = 1; i < limit; i++) {
      const val = masterValues[i][masterIdx];
      if (typeof val === "number") {
        numCount++;
        if (val > 10958 && val < 73051) dateCount++;
      }
    }
    if (numCount > 0 && dateCount / numCount > 0.5) {
      dateColumnNames.add(colConfig.name);
    }
  });

  // ── 5. Build color map (value → color) from Master List ──────────────
  const EXCLUDED_COLORS = new Set([
    "#ffffff", "#add8e6",
    "#fc0019", "#ff0000", "#ff0d0d", "#ff1a1a", "#fe0000",
    "#ff2400", "#cc0000", "#ee0000", "#dd0000", "#e60000",
  ]);

  const outputColMasterIndices = outputColumns
    .map((c) => getColIndex(c.name))
    .filter((idx) => idx !== -1);

  // Sample colors from first N rows
  const colorSampleLimit = Math.min(masterValues.length, 200);
  const columnColorMaps = new Map<number, Map<string, string>>();

  for (let r = 1; r < colorSampleLimit; r++) {
    for (const cIdx of outputColMasterIndices) {
      const val = masterValues[r][cIdx];
      if (!val) continue;
      const cell = masterSheet.getCell(r, cIdx);
      const color = cell.getFormat().getFill().getColor();
      if (!color) continue;
      const normColor = color.toLowerCase();
      if (EXCLUDED_COLORS.has(normColor)) continue;
      if (!columnColorMaps.has(cIdx)) {
        columnColorMaps.set(cIdx, new Map<string, string>());
      }
      const colMap = columnColorMaps.get(cIdx)!;
      if (!colMap.has(String(val))) {
        colMap.set(String(val), color);
      }
    }
  }

  // ── 6. Filter by Days Out ────────────────────────────────────────────
  interface RowObj {
    values: (string | number | boolean)[];
    formulas: string[];
    originalIndex: number;
  }

  const shouldDeferToFailing = DAYS_OUT_THRESHOLD < 5 && INCLUDE_FAILING && gradeIdx !== -1;

  const dataRows: RowObj[] = [];
  for (let i = 1; i < masterValues.length; i++) {
    const daysOutVal = masterValues[i][daysOutIdx];
    if (typeof daysOutVal === "number" && daysOutVal >= DAYS_OUT_THRESHOLD) {
      if (shouldDeferToFailing && daysOutVal < 5) {
        const gradeVal = masterValues[i][gradeIdx];
        const isFailing =
          typeof gradeVal === "number" && (gradeVal < 0.6 || (gradeVal >= 1 && gradeVal < 60));
        if (isFailing) continue;
      }
      dataRows.push({ values: masterValues[i], formulas: masterFormulas[i], originalIndex: i });
    }
  }
  dataRows.sort((a, b) => (Number(b.values[daysOutIdx]) || 0) - (Number(a.values[daysOutIdx]) || 0));

  // ── 7. Filter by Grade (Failing) ────────────────────────────────────
  let failingRows: RowObj[] = [];
  if (INCLUDE_FAILING && gradeIdx !== -1) {
    for (let i = 1; i < masterValues.length; i++) {
      const gradeVal = masterValues[i][gradeIdx];
      const daysOutVal = masterValues[i][daysOutIdx];
      const isFailing =
        typeof gradeVal === "number" && (gradeVal < 0.6 || (gradeVal >= 1 && gradeVal < 60));
      const isRecent = typeof daysOutVal === "number" && daysOutVal <= 4;
      if (isFailing && isRecent) {
        failingRows.push({ values: masterValues[i], formulas: masterFormulas[i], originalIndex: i });
      }
    }
    failingRows.sort((a, b) => (Number(a.values[gradeIdx]) || 0) - (Number(b.values[gradeIdx]) || 0));
  }

  // ── 8. Filter by Attendance ──────────────────────────────────────────
  let attendanceRows: RowObj[] = [];
  if (INCLUDE_ATTENDANCE && attendanceIdx !== -1) {
    const ldaStudentIds = new Set<string | number | boolean>();
    if (studentIdIdx !== -1) {
      for (const row of dataRows) {
        const sId = row.values[studentIdIdx];
        if (sId) ldaStudentIds.add(sId);
      }
    }
    const attendanceStudentIds = new Set<string | number | boolean>();

    for (let i = 1; i < masterValues.length; i++) {
      const sId = studentIdIdx !== -1 ? masterValues[i][studentIdIdx] : null;
      if (sId && ldaStudentIds.has(sId)) continue;

      const attVal = masterValues[i][attendanceIdx];
      let attPercent: number | null = null;
      if (typeof attVal === "number") {
        attPercent = attVal <= 1 ? attVal * 100 : attVal;
      }
      if (attPercent !== null && attPercent < 60) {
        attendanceRows.push({
          values: masterValues[i],
          formulas: masterFormulas[i],
          originalIndex: i,
        });
        if (sId) attendanceStudentIds.add(sId);
      }
    }
    attendanceRows.sort(
      (a, b) => {
        const aVal = typeof a.values[attendanceIdx] === "number" ? a.values[attendanceIdx] as number : 0;
        const bVal = typeof b.values[attendanceIdx] === "number" ? b.values[attendanceIdx] as number : 0;
        return (aVal <= 1 ? aVal * 100 : aVal) - (bVal <= 1 ? bVal * 100 : bVal);
      }
    );

    // Remove attendance students from failing list (no duplicates)
    if (attendanceStudentIds.size > 0 && studentIdIdx !== -1) {
      failingRows = failingRows.filter((row) => {
        const sId = row.values[studentIdIdx];
        return !sId || !attendanceStudentIds.has(sId);
      });
    }
  }

  // ── 9. Build output row ──────────────────────────────────────────────
  interface ProcessedRow {
    cells: (string | number | boolean)[];
    formulas: (string | null)[];
    cellHighlights: { colIndex: number; color: string }[];
  }

  function buildOutputRow(rowObj: RowObj): ProcessedRow {
    const cells: (string | number | boolean)[] = [];
    const formulas: (string | null)[] = [];
    const cellHighlights: { colIndex: number; color: string }[] = [];

    outputColumns.forEach((colConfig, colOutIdx) => {
      const masterIdx = getColIndex(colConfig.name);
      let val: string | number | boolean = masterIdx !== -1 ? rowObj.values[masterIdx] : "";
      let form: string | null = masterIdx !== -1 ? rowObj.formulas[masterIdx] : null;

      // Preserve Gradebook HYPERLINK formulas
      if (stripStr(colConfig.name) === "gradebook" && form && String(form).startsWith("=HYPERLINK")) {
        // Keep formula as-is
      } else if (stripStr(colConfig.name) === "gradebook" && val && String(val).startsWith("http")) {
        form = `=HYPERLINK("${val}", "Link")`;
        val = "Link";
      } else {
        form = null;
      }

      // Value-based color mapping (advisor colors, etc.)
      if (masterIdx !== -1 && val) {
        const colMap = columnColorMaps.get(masterIdx);
        if (colMap && colMap.has(String(val))) {
          cellHighlights.push({ colIndex: colOutIdx, color: colMap.get(String(val))! });
        }
      }

      cells.push(val);
      formulas.push(form);
    });

    return { cells, formulas, cellHighlights };
  }

  // ── 10. Detect campuses ──────────────────────────────────────────────
  let campusList: string[] = [];
  const isMultiCampus = SHEET_MODE === "campus" && campusIdx !== -1;

  if (isMultiCampus) {
    const campusSet = new Set<string>();
    for (let i = 1; i < masterValues.length; i++) {
      const val = String(masterValues[i][campusIdx] || "").trim();
      if (val) campusSet.add(val);
    }
    campusList = Array.from(campusSet).sort();
  }

  // If no campuses found or single mode, create one sheet
  if (campusList.length === 0) {
    if (SHEET_MODE === "campus" && campusIdx !== -1) {
      // Single campus — use its name
      for (let i = 1; i < masterValues.length; i++) {
        const val = String(masterValues[i][campusIdx] || "").trim();
        if (val) {
          campusList = [val];
          break;
        }
      }
    }
    if (campusList.length === 0) {
      const today = new Date();
      const dateStr = `${today.getMonth() + 1}-${today.getDate()}-${today.getFullYear()}`;
      campusList = [`LDA ${dateStr}`];
    }
  }

  // ── 11. Create sheets ────────────────────────────────────────────────
  const existingSheets = workbook.getWorksheets().map((s) => s.getName());
  let sheetsCreated = 0;
  let totalStudents = 0;

  for (let ci = 0; ci < campusList.length; ci++) {
    const campusName = campusList[ci];

    // Filter rows for this campus
    const filterByCampus = (rows: RowObj[]): RowObj[] => {
      if (!isMultiCampus || campusIdx === -1) return rows;
      return rows.filter((r) => String(r.values[campusIdx] || "").trim() === campusName);
    };

    const campusDataRows = filterByCampus(dataRows);
    const campusFailingRows = filterByCampus(failingRows);
    const campusAttendanceRows = filterByCampus(attendanceRows);

    // Generate unique sheet name
    let sheetName = campusName;
    let counter = 2;
    while (existingSheets.includes(sheetName)) {
      sheetName = `${campusName} (${counter++})`;
    }
    existingSheets.push(sheetName);

    const newSheet = workbook.addWorksheet(sheetName);
    if (ci === 0) newSheet.activate();

    // ── Write LDA table ────────────────────────────────────────────
    let currentRow = 0;
    if (campusDataRows.length > 0) {
      const processed = campusDataRows.map((r) => buildOutputRow(r));
      writeTable(newSheet, masterSheet, currentRow, `LDA_${ci}`, outputColumns, processed, dateColumnNames, getColIndex);
      currentRow = campusDataRows.length + 4;
      totalStudents += campusDataRows.length;
    } else {
      // Write headers even if empty
      const headerRange = newSheet.getRangeByIndexes(0, 0, 1, outputColumns.length);
      headerRange.setValues([outputColumns.map((c) => c.name)]);
      headerRange.getFormat().getFont().setBold(true);
      currentRow = 4;
    }

    // ── Write Failing table ────────────────────────────────────────
    if (INCLUDE_FAILING && campusFailingRows.length > 0) {
      const titleCell = newSheet.getRangeByIndexes(currentRow - 1, 0, 1, 1);
      titleCell.setValue("Failing Students (Active)");
      titleCell.getFormat().getFont().setBold(true);

      const processed = campusFailingRows.map((r) => buildOutputRow(r));
      writeTable(newSheet, masterSheet, currentRow, `Failing_${ci}`, outputColumns, processed, dateColumnNames, getColIndex);
      currentRow += campusFailingRows.length + 4;
      totalStudents += campusFailingRows.length;
    }

    // ── Write Attendance table ─────────────────────────────────────
    if (INCLUDE_ATTENDANCE && campusAttendanceRows.length > 0) {
      const titleCell = newSheet.getRangeByIndexes(currentRow - 1, 0, 1, 1);
      titleCell.setValue("Low Attendance Students");
      titleCell.getFormat().getFont().setBold(true);

      const processed = campusAttendanceRows.map((r) => buildOutputRow(r));
      writeTable(newSheet, masterSheet, currentRow, `Attendance_${ci}`, outputColumns, processed, dateColumnNames, getColIndex);
      totalStudents += campusAttendanceRows.length;
    }

    // ── Auto-fit columns ───────────────────────────────────────────
    const usedRange = newSheet.getUsedRange();
    if (usedRange) {
      usedRange.getFormat().autofitColumns();
    }

    // ── Hide implicit columns ──────────────────────────────────────
    outputColumns.forEach((colConfig, idx) => {
      if (colConfig.hidden) {
        newSheet.getRangeByIndexes(0, idx, 1, 1).getEntireColumn().setColumnHidden(true);
      }
    });

    sheetsCreated++;
  }

  return `SUCCESS: Created ${sheetsCreated} LDA sheet(s) with ${totalStudents} total students across ${campusList.join(", ")}.`;
}

/**
 * Writes a data table to a sheet with formatting, conditional formatting,
 * and cell highlights. Creates an Excel Table with TableStyleLight9.
 */
function writeTable(
  sheet: ExcelScript.Worksheet,
  masterSheet: ExcelScript.Worksheet,
  startRow: number,
  tableName: string,
  outputColumns: { name: string; hidden?: boolean }[],
  processedRows: {
    cells: (string | number | boolean)[];
    formulas: (string | null)[];
    cellHighlights: { colIndex: number; color: string }[];
  }[],
  dateColumnNames: Set<string>,
  getColIndex: (name: string) => number
): void {
  if (processedRows.length === 0) return;

  const rowCount = processedRows.length;
  const colCount = outputColumns.length;
  const headers = outputColumns.map((c) => c.name);

  // Write headers
  const headerRange = sheet.getRangeByIndexes(startRow, 0, 1, colCount);
  headerRange.setValues([headers]);

  // Write data
  const allValues = processedRows.map((r) => r.cells);
  const dataRange = sheet.getRangeByIndexes(startRow + 1, 0, rowCount, colCount);
  dataRange.setValues(allValues);

  // Apply formulas where they exist
  for (let r = 0; r < processedRows.length; r++) {
    for (let c = 0; c < processedRows[r].formulas.length; c++) {
      const formula = processedRows[r].formulas[c];
      if (formula) {
        sheet.getCell(startRow + 1 + r, c).setFormula(formula);
      }
    }
  }

  // Create table
  const fullRange = sheet.getRangeByIndexes(startRow, 0, rowCount + 1, colCount);
  const table = sheet.addTable(fullRange, true);
  table.setName(tableName + "_" + Math.floor(Math.random() * 1000));
  table.setPredefinedTableStyle("TableStyleLight9");

  // Copy conditional formatting from Master List
  const masterHeaders = masterSheet.getUsedRange()!.getValues()[0] as (string | number | boolean)[];
  outputColumns.forEach((colConfig, idx) => {
    const masterIdx = getColIndex(colConfig.name);
    if (masterIdx === -1) return;

    const sourceCell = masterSheet.getCell(1, masterIdx);
    const targetCol = table.getColumnByName(colConfig.name);
    if (!targetCol) return;
    const targetRange = targetCol.getRangeBetweenHeaderAndTotal();

    // Copy formats (conditional formatting) from master, then clear the fill
    // so table style shows through
    targetRange.copyFrom(sourceCell, ExcelScript.RangeCopyType.formats, false, false);
    targetRange.getFormat().getFill().clear();

    // Apply date format if this is a date column
    if (dateColumnNames.has(colConfig.name)) {
      targetRange.setNumberFormatLocal("mm-dd-yy;@");
    }
  });

  // Apply cell highlights (advisor colors, value-based colors)
  for (let r = 0; r < processedRows.length; r++) {
    const highlights = processedRows[r].cellHighlights;
    if (highlights.length === 0) continue;

    // Merge consecutive highlights with same color
    const sorted = [...highlights].sort((a, b) => a.colIndex - b.colIndex);
    let current: { startCol: number; endCol: number; color: string } | null = null;

    for (const h of sorted) {
      if (current && current.color === h.color && current.endCol === h.colIndex - 1) {
        current.endCol = h.colIndex;
      } else {
        if (current) {
          const range = sheet.getRangeByIndexes(
            startRow + 1 + r,
            current.startCol,
            1,
            current.endCol - current.startCol + 1
          );
          range.getFormat().getFill().setColor(current.color);
        }
        current = { startCol: h.colIndex, endCol: h.colIndex, color: h.color };
      }
    }
    if (current) {
      const range = sheet.getRangeByIndexes(
        startRow + 1 + r,
        current.startCol,
        1,
        current.endCol - current.startCol + 1
      );
      range.getFormat().getFill().setColor(current.color);
    }
  }
}
