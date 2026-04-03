/**
 * HighlightRows – Office Script for Power Automate
 *
 * Replicates the Student Retention Add-in's auto-highlighting logic on any
 * specified sheet. Reads the target sheet and optionally "Student History",
 * then applies all the same highlighting rules the add-in uses:
 *
 *   1. Retention status highlighting (partial row: Student Name → Outreach)
 *      - Orange (#FFEDD5): default retention message
 *      - Red (#FFC7CE): Do Not Contact (DNC)
 *      - Light green (#E2EFDA): next assignment is due (0 missing)
 *   2. Advisor / Assigned column color preservation (value → color map)
 *   3. DNC phone strikethrough (#FFC7CE + strikethrough dark red #9C0006)
 *   4. New-student light blue rows (#ADD8E6) based on latest Expected Start Date
 *   5. Conditional formatting:
 *      - Grade / Last Course Grade / Attendance: Red-Yellow-Green 3-color scale
 *      - Missing Assignments = 0: light green (#E2EFDA)
 *      - Hold = "Yes": light red (#FFB6C1)
 *      - AdSAPStatus contains "Financial": light red (#FFB6C1)
 *
 * @param sheetName - Name of the sheet to highlight. Defaults to "Master List".
 */
function main(workbook: ExcelScript.Workbook, sheetName?: string): string {
  // Default to "Master List" if no sheet name provided
  const targetSheetName = sheetName && sheetName.trim() !== "" ? sheetName.trim() : "Master List";
  // ── Constants ────────────────────────────────────────────────────────
  const STUDENT_NAME_COLS = ["student name", "studentname", "student"];
  const STUDENT_ID_COLS  = ["student id", "systudentid", "id"];
  const OUTREACH_COLS    = ["outreach"];
  const ASSIGNED_COLS    = ["assigned", "advisor"];
  const GRADEBOOK_COLS   = ["grade book", "gradebook", "gradebooklink", "gradelink"];
  const GRADE_COLS       = ["grade", "current score", "course grade", "grades"];
  const LAST_GRADE_COLS  = ["last course grade", "lastcoursegrade"];
  const MISSING_COLS     = ["missing assignments", "missingassignments", "course missing assignments", "total missing"];
  const HOLD_COLS        = ["hold"];
  const ADSAP_COLS       = ["adsapstatus"];
  const ATTENDANCE_COLS  = ["attendance %", "attendance%", "attendancepercent", "attendance"];
  const EXP_START_COLS   = ["expected start date", "start date", "expstartdate"];
  const DAYS_OUT_COLS    = ["days out"];
  const PHONE_COLS       = ["phone", "phone number", "phonenumber", "contact number"];
  const OTHER_PHONE_COLS = ["other phone", "otherphone"];
  const NEXT_DUE_COLS    = ["next assignment due", "nextassignmentdue"];

  // Colors excluded from color-map sampling (white, light blue new-student, reds)
  const EXCLUDED_COLORS = new Set([
    "#ffffff", "#add8e6",
    "#fc0019", "#ff0000", "#ff0d0d", "#ff1a1a", "#fe0000",
    "#ff2400", "#cc0000", "#ee0000", "#dd0000", "#e60000",
    "#ffff00" // also exclude yellow (contacted highlight)
  ]);

  // ── Helpers ──────────────────────────────────────────────────────────
  const normalize = (h: string) => h.toLowerCase().replace(/\s+/g, "");

  function findCol(headers: string[], aliases: string[]): number {
    const lower = headers.map((h) => h.toLowerCase());
    for (const a of aliases) {
      const i = lower.indexOf(a);
      if (i !== -1) return i;
    }
    const norm = headers.map(normalize);
    for (const a of aliases) {
      const i = norm.indexOf(normalize(a));
      if (i !== -1) return i;
    }
    return -1;
  }

  function parseDate(v: string | number | boolean): Date | null {
    if (v == null || v === "") return null;
    if (typeof v === "number" && v > 25569) {
      return new Date((v - 25569) * 86400 * 1000);
    }
    if (typeof v === "string") {
      const d = new Date(v);
      if (!isNaN(d.getTime())) return d;
    }
    return null;
  }

  function formatFriendlyDate(v: string | number | boolean): string {
    const dayNames = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"];
    const monthNames = ["January", "February", "March", "April", "May", "June",
      "July", "August", "September", "October", "November", "December"];
    let target: Date | null = null;
    if (typeof v === "number") {
      const d = new Date(Math.round((v - 25569) * 86400 * 1000));
      target = new Date(d.valueOf() + d.getTimezoneOffset() * 60000);
    } else if (typeof v === "string") {
      const match = v.trim().match(/^(\d{2})-(\d{2})-(\d{2})$/);
      if (match) {
        target = new Date(2000 + parseInt(match[3], 10), parseInt(match[1], 10) - 1, parseInt(match[2], 10));
      } else {
        const d = new Date(v);
        if (!isNaN(d.getTime())) target = d;
      }
    }
    if (!target) return String(v);
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    target.setHours(0, 0, 0, 0);
    const diff = Math.round((target.getTime() - today.getTime()) / 86400000);
    if (diff === 0) return "today";
    if (diff === 1) return "tomorrow";
    if (diff >= 2 && diff <= 6) return `this ${dayNames[target.getDay()]}`;
    if (diff >= 7 && diff <= 13) return `next ${dayNames[target.getDay()]}`;
    const day = target.getDate();
    const suf = (day === 1 || day === 21 || day === 31) ? "st"
      : (day === 2 || day === 22) ? "nd"
      : (day === 3 || day === 23) ? "rd" : "th";
    return `${monthNames[target.getMonth()]} ${day}${suf}`;
  }

  // ── 1. Get target sheet ───────────────────────────────────────────────
  const sheet = workbook.getWorksheet(targetSheetName);
  if (!sheet) return `ERROR: No '${targetSheetName}' sheet found.`;

  const usedRange = sheet.getUsedRange();
  if (!usedRange) return `ERROR: '${targetSheetName}' sheet is empty.`;

  const values = usedRange.getValues() as (string | number | boolean)[][];
  if (values.length < 2) return `ERROR: '${targetSheetName}' has no data rows.`;

  const headers = values[0].map((h) => String(h ?? ""));

  // ── 2. Identify columns ──────────────────────────────────────────────
  const nameCol      = findCol(headers, STUDENT_NAME_COLS);
  const idCol        = findCol(headers, STUDENT_ID_COLS);
  const outreachCol  = findCol(headers, OUTREACH_COLS);
  const assignedCol  = findCol(headers, ASSIGNED_COLS);
  const gradebookCol = findCol(headers, GRADEBOOK_COLS);
  const gradeCol     = findCol(headers, GRADE_COLS);
  const lastGradeCol = findCol(headers, LAST_GRADE_COLS);
  const missingCol   = findCol(headers, MISSING_COLS);
  const holdCol      = findCol(headers, HOLD_COLS);
  const adsapCol     = findCol(headers, ADSAP_COLS);
  const attendCol    = findCol(headers, ATTENDANCE_COLS);
  const expStartCol  = findCol(headers, EXP_START_COLS);
  const daysOutCol   = findCol(headers, DAYS_OUT_COLS);
  const phoneCol     = findCol(headers, PHONE_COLS);
  const otherPhoneCol = findCol(headers, OTHER_PHONE_COLS);
  const nextDueCol   = findCol(headers, NEXT_DUE_COLS);

  if (nameCol === -1) return "ERROR: No Student Name column found.";

  const dataRowCount = values.length - 1;

  // ── 3. Read Student History for DNC & LDA tags ───────────────────────
  const dncMap = new Map<string, string>();           // studentId → "dnc, dnc - phone, ..."
  const ldaFollowUpMap = new Map<string, Date>();     // studentId → follow-up date

  const historySheet = workbook.getWorksheet("Student History");
  if (historySheet && idCol !== -1) {
    const hUsed = historySheet.getUsedRange();
    if (hUsed) {
      const hValues = hUsed.getValues() as (string | number | boolean)[][];
      if (hValues.length > 1) {
        const hHeaders = hValues[0].map((h) => String(h).toLowerCase().trim());
        const hIdIdx = hHeaders.findIndex((h) =>
          (h.includes("student") && h.includes("id")) || h.includes("number")
        );
        const hTagIdx = hHeaders.indexOf("tag");

        if (hIdIdx !== -1 && hTagIdx !== -1) {
          const todayTime = new Date();
          todayTime.setHours(0, 0, 0, 0);
          const ldaRegex = /\blda\b.*?(\d{1,2}[-\/]\d{1,2}[-\/]\d{2,4})/i;

          // Read from bottom up (most recent tags first)
          for (let i = hValues.length - 1; i > 0; i--) {
            const hid = String(hValues[i][hIdIdx] ?? "").trim();
            const htagRaw = String(hValues[i][hTagIdx] ?? "");
            const htagLower = htagRaw.toLowerCase().trim();

            if (!hid) continue;

            // DNC tags: accumulate all DNC entries per student
            if (htagLower.includes("dnc")) {
              dncMap.set(hid, dncMap.has(hid) ? dncMap.get(hid)! + ", " + htagLower : htagLower);
            }

            // LDA follow-up: keep only the first (most recent) entry per student
            if (!ldaFollowUpMap.has(hid)) {
              const match = htagRaw.match(ldaRegex);
              if (match) {
                const ldaDate = new Date(match[1]);
                if (!isNaN(ldaDate.getTime())) {
                  ldaDate.setHours(0, 0, 0, 0);
                  if (ldaDate.getTime() >= todayTime.getTime()) {
                    ldaFollowUpMap.set(hid, ldaDate);
                  }
                }
              }
            }
          }
        }
      }
    }
  }

  // ── 4. Build color map from existing Assigned column colors ──────────
  // Maps value → color for each column (like the LDA processor does)
  const valueColorMap = new Map<string, string>(); // assigned value → color

  if (assignedCol !== -1) {
    const seen = new Set<string>();
    for (let r = 1; r < values.length; r++) {
      const val = String(values[r][assignedCol] ?? "").trim();
      if (!val || seen.has(val)) continue;
      seen.add(val);
      const cell = sheet.getCell(r, assignedCol);
      const color = cell.getFormat().getFill().getColor();
      if (color && !EXCLUDED_COLORS.has(color.toLowerCase())) {
        valueColorMap.set(val, color);
      }
    }
  }

  // ── 5. Check if "Next Assignment Due" column is entirely blank ───────
  let nextDueColumnAllBlank = true;
  if (nextDueCol !== -1) {
    for (let r = 1; r < values.length; r++) {
      const v = values[r][nextDueCol];
      if (v != null && String(v).trim() !== "") {
        nextDueColumnAllBlank = false;
        break;
      }
    }
  }

  // ── 6. Detect new students by latest Expected Start Date ─────────────
  const newStudentRows = new Set<number>();
  if (expStartCol !== -1) {
    let latestDate: Date | null = null;
    for (let r = 1; r < values.length; r++) {
      const d = parseDate(values[r][expStartCol]);
      if (d && (!latestDate || d > latestDate)) latestDate = d;
    }
    if (latestDate) {
      const latestStr = latestDate.toDateString();
      for (let r = 1; r < values.length; r++) {
        const d = parseDate(values[r][expStartCol]);
        if (d && d.toDateString() === latestStr) newStudentRows.add(r);
      }
    }
  }

  // ── 7. Clear existing fills & strikethroughs on data rows ────────────
  // (Keep header formatting intact)
  if (dataRowCount > 0) {
    const dataRange = sheet.getRangeByIndexes(1, 0, dataRowCount, headers.length);
    dataRange.getFormat().getFill().clear();
    dataRange.getFormat().getFont().setStrikethrough(false);
    dataRange.getFormat().getFont().setColor("#000000");
  }

  // Clear existing conditional formats
  const fullRange = sheet.getUsedRange();
  if (fullRange) {
    fullRange.clearAllConditionalFormats();
  }

  // ── 8. Determine the partial-row highlight range (Name → Outreach) ──
  const partialStart = outreachCol !== -1
    ? Math.min(nameCol, outreachCol)
    : -1;
  const partialEnd = outreachCol !== -1
    ? Math.max(nameCol, outreachCol)
    : -1;

  // ── 9. Apply row-by-row highlighting ─────────────────────────────────
  let highlightedCount = 0;
  let dncCount = 0;
  let ldaCount = 0;
  let contactedCount = 0;

  for (let r = 1; r < values.length; r++) {
    const studentId = idCol !== -1 ? String(values[r][idCol] ?? "").trim() : "";
    const missingVal = missingCol !== -1 ? values[r][missingCol] : null;
    const nextDueVal = nextDueCol !== -1 ? values[r][nextDueCol] : null;

    // ── 9a. Generate retention message (same priority logic as add-in) ──
    let retentionMsg: string | null = null;

    // Priority 1: DNC
    if (studentId && dncMap.has(studentId)) {
      const dncTag = dncMap.get(studentId)!;
      const tags = dncTag.split(",").map((t) => t.trim());
      const hasExcludable = tags.some(
        (t) => t.includes("dnc") && t !== "dnc - phone" && t !== "dnc - other phone"
      );
      if (hasExcludable) {
        retentionMsg = "Do not contact";
      }
    }

    // Priority 2: LDA follow-up
    if (!retentionMsg && studentId && ldaFollowUpMap.has(studentId)) {
      const ldaDate = ldaFollowUpMap.get(studentId)!;
      const monthNames = ["January", "February", "March", "April", "May", "June",
        "July", "August", "September", "October", "November", "December"];
      const day = ldaDate.getDate();
      const suf = (day === 1 || day === 21 || day === 31) ? "st"
        : (day === 2 || day === 22) ? "nd"
        : (day === 3 || day === 23) ? "rd" : "th";
      retentionMsg = `[Follow up] Student will engage on ${monthNames[ldaDate.getMonth()]} ${day}${suf}`;
      ldaCount++;
    }

    // Priority 3: Zero missing + next assignment due
    if (!retentionMsg && typeof missingVal === "number" && missingVal === 0 && nextDueVal) {
      retentionMsg = `Student's next assignment is due ${formatFriendlyDate(nextDueVal)}.`;
    }

    // Priority 4: Zero missing but NO next due date (anomaly)
    if (!retentionMsg && typeof missingVal === "number" && missingVal === 0
        && !nextDueVal && !nextDueColumnAllBlank && nextDueCol !== -1) {
      retentionMsg = "Student has 0 missing assignments but they have no next assignment due date. Please check their Grade Book.";
    }

    const isGradeBookFlag = retentionMsg !== null && retentionMsg.includes("Please check their Grade Book");
    const isRetentionActive = retentionMsg !== null && !isGradeBookFlag;
    const isDNC = retentionMsg !== null && retentionMsg.includes("DNC");
    const isNextDue = retentionMsg !== null && retentionMsg.startsWith("Student's next assignment is due");

    if (isDNC) dncCount++;

    // ── 9b. Determine partial-row color ─────────────────────────────────
    let partialRowColor = "#FFEDD5"; // Orange default
    if (isDNC) {
      partialRowColor = "#FFC7CE"; // Red for DNC
    } else if (isNextDue) {
      partialRowColor = "#E2EFDA"; // Light green
    }

    // ── 9c. Apply retention highlight (partial row: Name → Outreach) ────
    if (isRetentionActive) {
      highlightedCount++;
      if (partialStart !== -1 && partialEnd !== -1) {
        const range = sheet.getRangeByIndexes(r, partialStart, 1, partialEnd - partialStart + 1);
        range.getFormat().getFill().setColor(partialRowColor);
      } else {
        // No Outreach column → highlight full row
        const range = sheet.getRangeByIndexes(r, 0, 1, headers.length);
        range.getFormat().getFill().setColor(partialRowColor);
      }
    }

    // ── 9d. Write retention message to Outreach column ──────────────────
    if (isRetentionActive && outreachCol !== -1) {
      const currentOutreach = String(values[r][outreachCol] ?? "").trim();
      // Only write if Outreach is empty (don't overwrite existing notes)
      if (!currentOutreach) {
        sheet.getCell(r, outreachCol).setValue(retentionMsg);
      }
    }

    // ── 9e. Assigned column color (advisor color preservation) ──────────
    if (assignedCol !== -1) {
      const assignedVal = String(values[r][assignedCol] ?? "").trim();
      if (assignedVal && valueColorMap.has(assignedVal)) {
        sheet.getCell(r, assignedCol).getFormat().getFill().setColor(valueColorMap.get(assignedVal)!);
      }
    }

    // ── 9f. DNC phone strikethrough ─────────────────────────────────────
    if (studentId && dncMap.has(studentId)) {
      const dncTags = dncMap.get(studentId)!.split(",").map((t) => t.trim());
      const hasGeneral = dncTags.some((t) => t === "dnc");
      const hasPhone = dncTags.some((t) => t === "dnc - phone");
      const hasOtherPhone = dncTags.some((t) => t === "dnc - other phone");

      if (phoneCol !== -1 && (hasGeneral || hasPhone)) {
        const cell = sheet.getCell(r, phoneCol);
        cell.getFormat().getFill().setColor("#FFC7CE");
        cell.getFormat().getFont().setStrikethrough(true);
        cell.getFormat().getFont().setColor("#9C0006");
      }
      if (otherPhoneCol !== -1 && (hasGeneral || hasOtherPhone)) {
        const cell = sheet.getCell(r, otherPhoneCol);
        cell.getFormat().getFill().setColor("#FFC7CE");
        cell.getFormat().getFont().setStrikethrough(true);
        cell.getFormat().getFont().setColor("#9C0006");
      }
    }

    // ── 9g. New student highlighting (light blue) ───────────────────────
    // Applied AFTER retention so it layers underneath for non-retention students
    if (newStudentRows.has(r) && !isRetentionActive) {
      const range = sheet.getRangeByIndexes(r, 0, 1, headers.length);
      range.getFormat().getFill().setColor("#ADD8E6");
    }
  }

  // ── 10. Conditional formatting ───────────────────────────────────────

  // Grade: Red → Yellow → Green 3-color scale
  applyThreeColorScale(sheet, gradeCol, dataRowCount, values);

  // Last Course Grade: same scale
  applyThreeColorScale(sheet, lastGradeCol, dataRowCount, values);

  // Attendance: same scale
  applyThreeColorScale(sheet, attendCol, dataRowCount, values);

  // Missing Assignments = 0 → light green
  if (missingCol !== -1 && dataRowCount > 0) {
    const range = sheet.getRangeByIndexes(1, missingCol, dataRowCount, 1);
    const cf = range.addConditionalFormat(ExcelScript.ConditionalFormatType.cellValue);
    cf.getCellValue().getFormat().getFill().setColor("#E2EFDA");
    cf.getCellValue().setRule({
      formula1: "0",
      operator: ExcelScript.ConditionalCellValueOperator.equalTo,
    });
  }

  // Hold = "Yes" → light red
  if (holdCol !== -1 && dataRowCount > 0) {
    const range = sheet.getRangeByIndexes(1, holdCol, dataRowCount, 1);
    const cf = range.addConditionalFormat(ExcelScript.ConditionalFormatType.containsText);
    cf.getTextComparison().getFormat().getFill().setColor("#FFB6C1");
    cf.getTextComparison().setRule({
      operator: ExcelScript.ConditionalTextOperator.contains,
      text: "Yes",
    });
  }

  // AdSAPStatus contains "Financial" → light red
  if (adsapCol !== -1 && dataRowCount > 0) {
    const range = sheet.getRangeByIndexes(1, adsapCol, dataRowCount, 1);
    const cf = range.addConditionalFormat(ExcelScript.ConditionalFormatType.containsText);
    cf.getTextComparison().getFormat().getFill().setColor("#FFB6C1");
    cf.getTextComparison().setRule({
      operator: ExcelScript.ConditionalTextOperator.contains,
      text: "Financial",
    });
  }

  // ── 11. Auto-fit columns ─────────────────────────────────────────────
  sheet.getUsedRange()?.getFormat().autofitColumns();

  // ── 12. Summary ──────────────────────────────────────────────────────
  const parts: string[] = [
    `SUCCESS: Highlighted ${dataRowCount} rows on '${targetSheetName}'.`,
  ];
  if (highlightedCount > 0)
    parts.push(`${highlightedCount} retention highlights applied.`);
  if (dncCount > 0)
    parts.push(`${dncCount} DNC students marked.`);
  if (ldaCount > 0)
    parts.push(`${ldaCount} LDA follow-ups tagged.`);
  if (newStudentRows.size > 0)
    parts.push(`${newStudentRows.size} new students highlighted in blue.`);

  return parts.join(" ");
}

/**
 * Applies a Red → Yellow → Green 3-color scale to a numeric column.
 * Auto-detects 0-1 vs 0-100 scale.
 */
function applyThreeColorScale(
  sheet: ExcelScript.Worksheet,
  colIdx: number,
  rowCount: number,
  data: (string | number | boolean)[][]
): void {
  if (colIdx === -1 || rowCount === 0) return;

  const range = sheet.getRangeByIndexes(1, colIdx, rowCount, 1);

  // Detect scale
  let isPercent = false;
  for (let i = 1; i < Math.min(data.length, 11); i++) {
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
