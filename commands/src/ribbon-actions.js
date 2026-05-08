/*
 * ribbon-actions.js
 *
 * Handles ribbon button actions for the Student Retention Add-in.
 * Implements the "Contacted" (toggle highlight) and "Transfer Data" button functionality.
 */
import { CONSTANTS } from './constants.js';
import { findColumnIndex, normalizeHeader } from '../../shared/excel-helpers.js';

/**
 * Creates a sendToCallQueue ribbon action bound to the given chromeExtensionService.
 * Reads the currently selected row(s), extracts student data, and sends it
 * to the Chrome Extension call queue via SRK_SELECTED_STUDENTS.
 *
 * @param {Object} extensionService - The chromeExtensionService singleton
 * @returns {Function} The ribbon action handler
 */
export function createSendToCallQueue(extensionService) {
  return async function sendToCallQueue(event) {
    try {
      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const selectedRange = context.workbook.getSelectedRange();
        selectedRange.load("rowIndex, rowCount, columnIndex, columnCount, values");
        const usedRange = sheet.getUsedRange();
        usedRange.load(["values", "rowIndex", "columnIndex"]);

        await context.sync();

        const allValues = usedRange.values;
        const headers = allValues[0].map(normalizeHeader);

        // Check if the selected cell value is already a phone number
        const selRowCount = selectedRange.rowCount || 1;
        const selColCount = selectedRange.columnCount || 1;

        if (selRowCount === 1 && selColCount === 1) {
          const cellValue = String(selectedRange.values?.[0]?.[0] || '').trim();
          const cleaned = cellValue.replace(/[\s\-\(\)\.]/g, '');

          if (cellValue && /^\+?\d{7,15}$/.test(cleaned)) {
            // Cell is a phone number — send directly
            const relativeRow = selectedRange.rowIndex - usedRange.rowIndex;
            const cellColIndex = selectedRange.columnIndex - usedRange.columnIndex;
            const rowData = (relativeRow > 0 && relativeRow < allValues.length)
              ? allValues[relativeRow] : null;

            const nameColIndex = findColumnIndex(headers, CONSTANTS.STUDENT_NAME_COLS);
            const phoneColIndex = findColumnIndex(headers, CONSTANTS.COLUMN_MAPPINGS.primaryPhone);
            const otherPhoneColIndex = findColumnIndex(headers, CONSTANTS.COLUMN_MAPPINGS.otherPhone);

            const name = (rowData && nameColIndex !== -1) ? String(rowData[nameColIndex] || '') : '';

            // Determine which phone field this cell belongs to
            let phone = '';
            let otherPhone = '';

            if (cellColIndex === otherPhoneColIndex) {
              // Selected cell is in the "other phone" column
              otherPhone = cellValue;
              phone = (rowData && phoneColIndex !== -1) ? String(rowData[phoneColIndex] || '') : '';
            } else {
              // Selected cell is in the primary phone column (or unknown column)
              phone = cellValue;
              otherPhone = (rowData && otherPhoneColIndex !== -1) ? String(rowData[otherPhoneColIndex] || '') : '';
            }

            extensionService.sendSelectedStudents(
              [{ name, syStudentId: '', phone, otherPhone }],
              cellValue,
              true
            );
            console.log(`sendToCallQueue: Direct phone ${cellValue} sent to call queue.`);
            return;
          }
        }

        // Not a direct phone number — look up columns from headers
        const nameColIndex = findColumnIndex(headers, CONSTANTS.STUDENT_NAME_COLS);
        const phoneColIndex = findColumnIndex(headers, CONSTANTS.COLUMN_MAPPINGS.primaryPhone);
        const otherPhoneColIndex = findColumnIndex(headers, CONSTANTS.COLUMN_MAPPINGS.otherPhone);

        if (phoneColIndex === -1 && otherPhoneColIndex === -1 && nameColIndex === -1) {
          console.error("sendToCallQueue: No phone or name columns found.");
          return;
        }

        // Detect if the selection falls within a specific phone column
        const selColStart = selectedRange.columnIndex - usedRange.columnIndex;
        const selColEnd = selColStart + (selectedRange.columnCount || 1) - 1;
        const selectedInOtherPhone = otherPhoneColIndex !== -1
          && selColStart <= otherPhoneColIndex && otherPhoneColIndex <= selColEnd;
        const selectedInPrimaryPhone = phoneColIndex !== -1
          && selColStart <= phoneColIndex && phoneColIndex <= selColEnd;

        const selectionStartRow = selectedRange.rowIndex;
        const selectionRowCount = selectedRange.rowCount;
        const dataStartRow = usedRange.rowIndex;

        // Load per-row hidden state so we can skip rows filtered out of view.
        const rowHiddenProxies = [];
        for (let i = 0; i < selectionRowCount; i++) {
          const rowRange = selectedRange.getRow(i).getEntireRow();
          rowRange.load("rowHidden");
          rowHiddenProxies.push(rowRange);
        }
        await context.sync();
        const rowHiddenStates = rowHiddenProxies.map(r => r.rowHidden === true);

        const students = [];
        const seenPhones = new Set(); // prevent duplicate entries for the same student

        for (let i = 0; i < selectionRowCount; i++) {
          const relativeRow = (selectionStartRow + i) - dataStartRow;

          // Skip header row and rows outside used range
          if (relativeRow <= 0 || relativeRow >= allValues.length) continue;

          // Skip rows that are hidden (filtered out or manually hidden)
          if (rowHiddenStates[i]) continue;

          const rowData = allValues[relativeRow];

          const name = nameColIndex !== -1 ? String(rowData[nameColIndex] || '') : '';
          const phone = phoneColIndex !== -1 ? String(rowData[phoneColIndex] || '') : '';
          const otherPhone = otherPhoneColIndex !== -1 ? String(rowData[otherPhoneColIndex] || '') : '';

          // Determine which number to dial based on selected column
          let directPhone = '';
          let isOtherContact = false;

          if (selectedInOtherPhone && !selectedInPrimaryPhone && otherPhone) {
            // Selection is only in the other phone column
            directPhone = otherPhone;
            isOtherContact = true;
          } else if (selectedInPrimaryPhone && !selectedInOtherPhone && phone) {
            directPhone = phone;
          }
          // If selection spans both columns or neither, no directPhone override

          // Skip rows with no phone number to dial
          const dialNumber = directPhone || phone || otherPhone;
          if (!dialNumber) continue;

          // Deduplicate: don't queue the same phone number twice
          // (prevents primary + other phone of the same student both being queued)
          const normalizedDial = dialNumber.replace(/\D/g, '');
          if (seenPhones.has(normalizedDial)) {
            console.log(`sendToCallQueue: Skipping duplicate phone ${dialNumber} for "${name}"`);
            continue;
          }
          seenPhones.add(normalizedDial);

          // Also skip if this student's OTHER number is already queued
          // (prevents same student appearing twice with different numbers)
          const counterpart = isOtherContact ? phone : otherPhone;
          if (counterpart) {
            const normalizedCounterpart = counterpart.replace(/\D/g, '');
            if (seenPhones.has(normalizedCounterpart)) {
              console.log(`sendToCallQueue: Skipping "${name}" — already queued with other number`);
              continue;
            }
          }

          students.push({
            name, syStudentId: '', phone, otherPhone,
            ...(directPhone ? { directPhone } : {}),
            ...(isOtherContact ? { isOtherContact: true } : {})
          });
        }

        if (students.length === 0) {
          console.log("sendToCallQueue: No valid student rows in selection.");
          return;
        }

        extensionService.sendSelectedStudents(students, null, true);
        console.log(`sendToCallQueue: Sent ${students.length} student(s) to call queue (autoCall).`);
      });
    } catch (error) {
      console.error("Error in sendToCallQueue: " + error);
      if (error instanceof OfficeExtension.Error) {
        console.error("Debug info: " + JSON.stringify(error.debugInfo));
      }
    } finally {
      if (event) {
        event.completed();
      }
    }
  };
}

/**
 * Finds the row of the current selection, and toggles a yellow highlight on the cells
 * in that row between the "StudentName" and "Outreach" columns.
 */
export async function toggleHighlight(event) {
  try {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const selectedRange = context.workbook.getSelectedRange();
      selectedRange.load("rowIndex");
      const usedRange = sheet.getUsedRange();
      usedRange.load(["rowIndex", "values"]);
      
      await context.sync();

      const headers = usedRange.values[0];
      const normalizedHeaders = headers.map(normalizeHeader);
      const studentNameColIndex = findColumnIndex(normalizedHeaders, CONSTANTS.STUDENT_NAME_COLS);
      const outreachColIndex = findColumnIndex(normalizedHeaders, CONSTANTS.OUTREACH_COLS);

      if (studentNameColIndex === -1 || outreachColIndex === -1) {
        console.error("Could not find 'StudentName' and/or 'Outreach' columns.");
        return; 
      }
      
      const startCol = Math.min(studentNameColIndex, outreachColIndex);
      const endCol = Math.max(studentNameColIndex, outreachColIndex);
      const colCount = endCol - startCol + 1;
      const targetRowIndex = selectedRange.rowIndex;

      if (targetRowIndex < usedRange.rowIndex) return;

      const highlightRange = sheet.getRangeByIndexes(targetRowIndex, startCol, 1, colCount);
      highlightRange.format.fill.load("color");
      
      await context.sync();

      if (highlightRange.format.fill.color === "#FFFF00") {
        highlightRange.format.fill.clear();
      } else {
        highlightRange.format.fill.color = "yellow";
      }

      await context.sync();
    });
  } catch (error) {
    console.error("Error in toggleHighlight: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.error("Debug info: " + JSON.stringify(error.debugInfo));
    }
  } finally {
    if (event) {
      event.completed();
    }
  }
}
