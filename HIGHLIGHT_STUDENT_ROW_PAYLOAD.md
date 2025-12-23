# Student Row Highlight Payload Guide

This document describes how to send a message from the Chrome Extension to the Excel Add-in to highlight a student row in the background.

## Overview

The Excel Add-in listens for `SRK_HIGHLIGHT_STUDENT_ROW` messages via `window.postMessage()`. When received, it will automatically highlight the specified student's row in the target worksheet **without requiring the side panel to be open**.

## Message Type

```
SRK_HIGHLIGHT_STUDENT_ROW
```

## Complete Payload Structure

```javascript
{
  type: "SRK_HIGHLIGHT_STUDENT_ROW",
  data: {
    studentName: string,      // Student's full name (for logging purposes)
    syStudentId: string,       // Student's SyStudentID (required - used to find row)
    startCol: number,          // Starting column index (0-based, required)
    endCol: number,            // Ending column index (0-based, required)
    targetSheet: string,       // Name of the worksheet (required)
    color: string,             // Hex color code (optional, defaults to #FFFF00 yellow)
    editColumn: number,        // Column index to edit (0-based, optional)
    editText: string           // Text to set in the edit column (optional)
  }
}
```

## Parameter Details

| Parameter | Type | Required | Description | Example |
|-----------|------|----------|-------------|---------|
| `type` | string | **Yes** | Message type identifier | `"SRK_HIGHLIGHT_STUDENT_ROW"` |
| `data.studentName` | string | No | Student's full name (used for logging) | `"John Doe"` |
| `data.syStudentId` | string | **Yes** | Student's unique identifier | `"12345678"` |
| `data.startCol` | number | **Yes** | 0-based starting column index | `0` (Column A) |
| `data.endCol` | number | **Yes** | 0-based ending column index | `5` (Column F) |
| `data.targetSheet` | string | **Yes** | Exact name of the worksheet | `"Fall 2024"` |
| `data.color` | string | No | Hex color code for highlight | `"#FF6B6B"` (red) |
| `data.editColumn` | number | No | Column index to edit (0-based) | `8` (Column I) |
| `data.editText` | string | No | Text to write to the cell | `"Submitted Midterm"` |

### Column Index Reference

Excel columns use 0-based indexing:

```
A = 0, B = 1, C = 2, D = 3, E = 4, F = 5, G = 6, H = 7, ...
```

**Examples:**
- Highlight columns A-F: `startCol: 0, endCol: 5`
- Highlight columns C-E: `startCol: 2, endCol: 4`
- Highlight only column B: `startCol: 1, endCol: 1`

### Color Format

Colors must be in **hexadecimal format** with `#` prefix:

```javascript
// Valid colors
"#FFFF00"  // Yellow (default)
"#FF6B6B"  // Red
"#4ECDC4"  // Teal
"#95E1D3"  // Mint
"#FFA07A"  // Light Salmon

// Invalid (will cause errors)
"yellow"   // ❌ Color names not supported
"rgb(255, 255, 0)"  // ❌ RGB format not supported
```

### Editing Cells (Optional)

In addition to highlighting, you can **edit a specific cell** in the student's row by providing `editColumn` and `editText`.

**Use Cases:**
- Track assignment submissions: "Submitted Midterm"
- Update outreach status: "Email Sent"
- Mark attendance: "Present"
- Record completion: "Completed Quiz 1"

**How it works:**
- Both `editColumn` (number) and `editText` (string) must be provided
- `editColumn` uses the same 0-based indexing as columns (A=0, B=1, etc.)
- The cell at the intersection of the student's row and `editColumn` will be updated with `editText`
- Cell editing happens **after** highlighting in the same operation

**Examples:**

```javascript
// Assuming column I (index 8) is the "Outreach" column
editColumn: 8,
editText: "Submitted Final Exam"

// Column F (index 5) for "Status"
editColumn: 5,
editText: "Email Sent"

// Column C (index 2) for "Attendance"
editColumn: 2,
editText: "Present"
```

## How to Send from Chrome Extension

### Method 1: Using Content Script (Recommended)

If your extension has a content script running in the Excel Online page:

```javascript
// From your content script
function highlightStudentRow(studentId, studentName, sheetName) {
  const message = {
    type: "SRK_HIGHLIGHT_STUDENT_ROW",
    data: {
      studentName: studentName,
      syStudentId: studentId,
      startCol: 0,        // Highlight from column A
      endCol: 8,          // to column I
      targetSheet: sheetName,
      color: "#FFFF00"    // Yellow highlight
    }
  };

  // Send to the page
  window.postMessage(message, "*");
}

// Usage example
highlightStudentRow("12345678", "John Doe", "Fall 2024");
```

### Method 2: Using Background Script with Tab Messaging

If sending from a background/service worker:

```javascript
// From background.js or service_worker.js
async function highlightStudentInExcel(tabId, studentData) {
  const message = {
    type: "SRK_HIGHLIGHT_STUDENT_ROW",
    data: {
      studentName: studentData.name,
      syStudentId: studentData.id,
      startCol: 0,
      endCol: 8,
      targetSheet: studentData.sheetName,
      color: "#4ECDC4"  // Teal highlight
    }
  };

  // Send to content script, which forwards to page
  await chrome.tabs.sendMessage(tabId, {
    action: "postToPage",
    message: message
  });
}
```

Then in your content script:

```javascript
// content.js - Relay messages from background to page
chrome.runtime.onMessage.addListener((request, sender, sendResponse) => {
  if (request.action === "postToPage") {
    window.postMessage(request.message, "*");
    sendResponse({ success: true });
  }
});
```

## Complete Working Examples

### Example 1: Highlight Single Student (Full Row)

```javascript
window.postMessage({
  type: "SRK_HIGHLIGHT_STUDENT_ROW",
  data: {
    studentName: "Jane Smith",
    syStudentId: "87654321",
    startCol: 0,
    endCol: 10,
    targetSheet: "Spring 2025",
    color: "#FFFF00"
  }
}, "*");
```

**Result:** Highlights columns A through K in Jane Smith's row with yellow.

### Example 2: Highlight Specific Columns (Grade Range)

```javascript
// Highlight only grade columns (assuming columns E-H contain grades)
window.postMessage({
  type: "SRK_HIGHLIGHT_STUDENT_ROW",
  data: {
    studentName: "Bob Johnson",
    syStudentId: "11223344",
    startCol: 4,   // Column E
    endCol: 7,     // Column H
    targetSheet: "Fall 2024",
    color: "#FF6B6B"  // Red for attention
  }
}, "*");
```

**Result:** Highlights only columns E-H in Bob Johnson's row with red.

### Example 3: Assignment Submission (Highlight + Edit Cell)

```javascript
// When a student submits an assignment, highlight their row green
// and update the "Outreach" column with submission info
window.postMessage({
  type: "SRK_HIGHLIGHT_STUDENT_ROW",
  data: {
    studentName: "Sarah Williams",
    syStudentId: "99887766",
    startCol: 0,
    endCol: 10,
    targetSheet: "Fall 2024",
    color: "#90EE90",  // Light green for success
    editColumn: 8,     // Column I (Outreach column)
    editText: "Submitted Midterm Exam"
  }
}, "*");
```

**Result:** Highlights Sarah Williams' row (columns A-K) in light green AND updates column I with "Submitted Midterm Exam".

### Example 4: Minimal Required Payload

```javascript
// Only required fields
window.postMessage({
  type: "SRK_HIGHLIGHT_STUDENT_ROW",
  data: {
    syStudentId: "55667788",
    startCol: 0,
    endCol: 5,
    targetSheet: "Winter 2025"
    // studentName omitted - won't break anything
    // color omitted - defaults to yellow (#FFFF00)
  }
}, "*");
```

### Example 5: Batch Highlighting Multiple Students

```javascript
// Highlight multiple students with different colors
const studentsToHighlight = [
  { id: "12345678", name: "Student A", color: "#FF6B6B" },
  { id: "23456789", name: "Student B", color: "#4ECDC4" },
  { id: "34567890", name: "Student C", color: "#95E1D3" }
];

studentsToHighlight.forEach(student => {
  window.postMessage({
    type: "SRK_HIGHLIGHT_STUDENT_ROW",
    data: {
      studentName: student.name,
      syStudentId: student.id,
      startCol: 0,
      endCol: 10,
      targetSheet: "Fall 2024",
      color: student.color
    }
  }, "*");
});
```

## Validation & Error Handling

The Add-in performs the following validations:

### Required Field Validation
- ✅ Checks that `syStudentId` is present
- ✅ Checks that `targetSheet` is present
- ✅ Checks that `startCol` and `endCol` are numbers
- ✅ Verifies Excel API is available

### Range Validation
- ✅ Ensures `startCol >= 0`
- ✅ Ensures `endCol >= 0`
- ✅ Ensures `startCol <= endCol`

### Runtime Validation
- ✅ Verifies the target sheet exists
- ✅ Searches for Student ID column using aliases
- ✅ Searches for matching student by ID

### Error Messages

The Add-in logs helpful error messages to the console:

```javascript
// Missing parameters
"ChromeExtensionService: Missing required parameters for highlight"

// Invalid column indices
"ChromeExtensionService: Invalid column indices"
"ChromeExtensionService: Invalid column range"

// Sheet not found
"ChromeExtensionService: Sheet 'SheetName' not found"

// Student ID column not found
"ChromeExtensionService: Could not find Student ID column in sheet"

// Student not found
"ChromeExtensionService: Student with ID '12345678' not found in sheet 'SheetName'"
```

## Listening for Results (Optional)

The Add-in emits events through its listener system when highlighting succeeds or fails.

### Subscribe to Events

```javascript
// In your content script, add listener to chromeExtensionService
// (Note: This requires access to the Add-in's context)

// The service will notify listeners of completion/errors
chromeExtensionService.addListener((event) => {
  switch (event.type) {
    case 'highlight_complete':
      console.log('✅ Highlight successful:', event.data);
      // event.data contains: studentName, syStudentId, targetSheet, startCol, endCol, color, editColumn, editText, timestamp
      break;

    case 'highlight_error':
      console.error('❌ Highlight failed:', event.data);
      // event.data contains: studentName, syStudentId, editColumn, editText, error, timestamp
      break;
  }
});
```

### Success Event Structure

```javascript
{
  type: "highlight_complete",
  data: {
    studentName: "John Doe",
    syStudentId: "12345678",
    targetSheet: "Fall 2024",
    startCol: 0,
    endCol: 8,
    color: "#FFFF00",
    editColumn: 8,              // Present if cell was edited
    editText: "Submitted Quiz", // Present if cell was edited
    timestamp: "2025-12-23T10:30:00.000Z"
  }
}
```

### Error Event Structure

```javascript
{
  type: "highlight_error",
  data: {
    studentName: "John Doe",
    syStudentId: "12345678",
    editColumn: 8,              // Present if editing was attempted
    editText: "Submitted Quiz", // Present if editing was attempted
    error: "Sheet 'Fall 2024' not found",
    timestamp: "2025-12-23T10:30:00.000Z"
  }
}
```

## How It Works Internally

1. **Message Reception**: Add-in's `chromeExtensionService` listens for messages via `window.addEventListener("message")`

2. **Handler Invocation**: When `SRK_HIGHLIGHT_STUDENT_ROW` is received, `handleHighlightStudentRow()` is called

3. **Sheet Lookup**: The function gets the worksheet by name using `getItemOrNullObject(targetSheet)`

4. **Student ID Search**:
   - Loads all values from the used range
   - Searches for Student ID column using aliases: `['Student ID', 'SyStudentID', 'Student identifier', 'ID']`
   - Finds the row where the ID column matches `syStudentId`

5. **Highlighting**: Uses `worksheet.getRangeByIndexes()` to select the specified column range and applies the fill color

6. **Cell Editing (Optional)**: If `editColumn` and `editText` are provided, updates the cell at the specified column in the student's row

7. **Notification**: Emits success or error events to any registered listeners

## Important Notes

- ✅ **No side panel required** - Works completely in the background
- ✅ **Automatic Student ID detection** - Supports multiple column name variations
- ✅ **Case-insensitive matching** - Student IDs are compared after trimming and converting to strings
- ✅ **Optional cell editing** - Can update a specific cell while highlighting the row
- ⚠️ **Exact sheet name required** - Sheet names must match exactly (case-sensitive)
- ⚠️ **Student must exist** - The student ID must be present in the specified sheet
- ⚠️ **Column indices are 0-based** - Column A = 0, not 1
- ⚠️ **Cell editing requires both parameters** - Both `editColumn` and `editText` must be provided to edit a cell

## Troubleshooting

### Highlight Not Appearing

1. **Check console for errors** - Open Excel Online DevTools and look for `ChromeExtensionService` logs
2. **Verify sheet name** - Ensure the exact sheet name is used (case-sensitive)
3. **Check student ID** - Verify the ID exists in the sheet and matches exactly
4. **Confirm column range** - Make sure `startCol` and `endCol` are valid indices

### Common Mistakes

```javascript
// ❌ WRONG - Using 1-based column indices
startCol: 1,  // This is column B, not A!

// ✅ CORRECT - Using 0-based indices
startCol: 0,  // Column A

// ❌ WRONG - Using color name
color: "red"

// ✅ CORRECT - Using hex code
color: "#FF0000"

// ❌ WRONG - Mismatched sheet name case
targetSheet: "fall 2024"  // Sheet is actually named "Fall 2024"

// ✅ CORRECT - Exact match
targetSheet: "Fall 2024"

// ❌ WRONG - Providing only editColumn without editText
editColumn: 8  // Missing editText!

// ✅ CORRECT - Both parameters provided
editColumn: 8,
editText: "Submitted Assignment"

// ❌ WRONG - Using 1-based index for editColumn
editColumn: 9  // This is column J, not I!

// ✅ CORRECT - Using 0-based index
editColumn: 8  // Column I
```

## Support

For issues or questions, refer to the main Student Retention Add-in documentation or check the console logs for detailed error messages.

---

**Last Updated**: December 2025
**Add-in Version**: 1.x
**Excel API Version**: 1.1+
