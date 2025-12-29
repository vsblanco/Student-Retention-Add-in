# Chrome Extension Integration Documentation

## Overview

The Student Retention Add-in communicates with the Chrome Extension using the `window.postMessage()` API. The background script (`background-services/background-service.js`) automatically syncs the Master List data when the extension is detected.

## Message Flow

### 1. Extension Detection

When the add-in loads, it pings to detect if the Chrome extension is installed:

**Add-in → Extension:**
```javascript
{
  type: "SRK_CHECK_EXTENSION"
}
```

**Extension → Add-in:**
```javascript
{
  type: "SRK_EXTENSION_INSTALLED"
}
```

### 2. Master List Sync Request

When the extension is detected, the add-in asks if it wants the Master List data:

**Add-in → Extension:**
```javascript
{
  type: "SRK_REQUEST_MASTER_LIST",
  timestamp: "2025-12-22T10:30:00.000Z"
}
```

**Extension → Add-in:**
```javascript
{
  type: "SRK_MASTER_LIST_RESPONSE",
  wantsData: true  // or false
}
```

### 3. Master List Data Transfer

If the extension responds with `wantsData: true`, the add-in sends the full Master List:

**Add-in → Extension:**
```javascript
{
  type: "SRK_MASTER_LIST_DATA",
  data: {
    sheetName: "Master List",
    columnMapping: { /* see below */ },
    students: [ /* array of student objects */ ],
    totalStudents: 150,
    timestamp: "2025-12-22T10:30:01.500Z"
  }
}
```

### 4. Import Master List from Extension

The Chrome extension can send master list data to import into the add-in's Master List sheet. The add-in will check if a Master List sheet exists before importing.

**Extension → Add-in:**
```javascript
{
  type: "SRK_IMPORT_MASTER_LIST",
  data: {
    headers: ["StudentName", "Student ID", "Grade", "Phone", ...],
    data: [
      ["Smith, John", "12345", 85.5, "555-1234", ...],
      ["Doe, Jane", "67890", 72, "555-9876", ...],
      // ... more rows
    ]
  }
}
```

**Import Behavior:**
- If the Master List sheet doesn't exist, the import will be aborted
- If the sheet exists, the add-in will:
  - Match incoming headers to Master List headers (case-insensitive with whitespace normalization)
  - Automatically wrap Gradebook URLs in HYPERLINK formulas with "Grade Book" as the display text
  - Preserve existing Gradebook hyperlinks for students already in the list
  - Preserve "Assigned" column values and their colors
  - Apply 3-color conditional formatting to Grade column (Red → Yellow → Green)
  - Highlight new students in light blue (#ADD8E6)
  - Clear and repopulate the sheet with the imported data

## Payload Structure

### Master List Data Payload (Add-in → Extension)

```typescript
interface MasterListPayload {
  type: "SRK_MASTER_LIST_DATA";
  data: {
    sheetName: string;                    // Always "Master List"
    headers: string[];                    // Array of all column headers
    columnMapping: Record<string, number>; // Map of header name -> column index
    students: Student[];                  // Array of student records (all columns included)
    totalStudents: number;                // Total count
    timestamp: string;                    // ISO 8601 timestamp
  }
}
```

### Import Master List Payload (Extension → Add-in)

```typescript
interface ImportMasterListPayload {
  type: "SRK_IMPORT_MASTER_LIST";
  data: {
    headers: string[];    // Array of column headers
    data: any[][];        // 2D array of row data (each row matches headers order)
  }
}
```

**Example:**
```json
{
  "type": "SRK_IMPORT_MASTER_LIST",
  "data": {
    "headers": ["StudentName", "Student ID", "Grade", "Phone", "Email"],
    "data": [
      ["Smith, John", "12345", 85.5, "555-1234", "john@school.edu"],
      ["Doe, Jane", "67890", 72, "555-9876", "jane@school.edu"]
    ]
  }
}

// The Student interface is dynamic - it includes ALL columns from your Master List
interface Student {
  [key: string]: any;  // Dynamic properties based on your Master List columns

  // Common fields (but not limited to these):
  // "Student Name"?: string;
  // "Student ID"?: string;
  // "Student Number"?: string;
  // "Gradebook"?: string;         // URL extracted from HYPERLINK formula
  // "Grade"?: number | string;
  // "Days Out"?: number;
  // "LDA"?: any;
  // "Phone"?: string;
  // "Other Phone"?: string;
  // "StudentEmail"?: string;
  // "PersonalEmail"?: string;
  // "Assigned"?: string;
  // "Outreach"?: string;
  // "Gender"?: string;
  // "Shift"?: string;
  // "ProgramVersion"?: string;
  // ... and any other custom columns in your sheet
}
```

## Example Payload

```json
{
  "type": "SRK_MASTER_LIST_DATA",
  "data": {
    "sheetName": "Master List",
    "headers": [
      "Assigned",
      "Student Name",
      "Student Number",
      "Gradebook",
      "Grade",
      "Missing Assignments",
      "LDA",
      "Days Out",
      "Gender",
      "Shift",
      "Outreach",
      "ProgramVersion",
      "SyStudentId",
      "Phone",
      "Other Phone",
      "StudentEmail",
      "PersonalEmail"
    ],
    "columnMapping": {
      "Assigned": 0,
      "Student Name": 1,
      "Student Number": 2,
      "Gradebook": 3,
      "Grade": 4,
      "Missing Assignments": 5,
      "LDA": 6,
      "Days Out": 7,
      "Gender": 8,
      "Shift": 9,
      "Outreach": 10,
      "ProgramVersion": 11,
      "SyStudentId": 12,
      "Phone": 13,
      "Other Phone": 14,
      "StudentEmail": 15,
      "PersonalEmail": 16
    },
    "students": [
      {
        "Assigned": "Dr. Johnson",
        "Student Name": "Smith, John",
        "Student Number": "STU001",
        "Gradebook": "https://nuc.instructure.com/courses/123/grades/12345",
        "Grade": 85.5,
        "Missing Assignments": 2,
        "LDA": "12/10/2025",
        "Days Out": 3,
        "Gender": "M",
        "Shift": "Day",
        "Outreach": "Called on 12/15",
        "ProgramVersion": "2.0",
        "SyStudentId": "12345",
        "Phone": "555-1234",
        "Other Phone": "555-5678",
        "StudentEmail": "john.smith@school.edu",
        "PersonalEmail": "john@email.com"
      },
      {
        "Assigned": "Dr. Williams",
        "Student Name": "Doe, Jane",
        "Student Number": "STU002",
        "Gradebook": "https://nuc.instructure.com/courses/123/grades/67890",
        "Grade": 72,
        "Missing Assignments": 5,
        "LDA": "12/05/2025",
        "Days Out": 7,
        "Gender": "F",
        "Shift": "Evening",
        "SyStudentId": "67890",
        "Phone": "555-9876",
        "StudentEmail": "jane.doe@school.edu"
      }
    ],
    "totalStudents": 2,
    "timestamp": "2025-12-22T10:30:01.500Z"
  }
}
```

## Chrome Extension Implementation Guide

### 1. Listen for Messages

```javascript
// In your Chrome extension content script or background script
window.addEventListener("message", (event) => {
  if (!event.data || !event.data.type) return;

  switch (event.data.type) {
    case "SRK_CHECK_EXTENSION":
      handleExtensionCheck();
      break;
    case "SRK_REQUEST_MASTER_LIST":
      handleMasterListRequest(event.data);
      break;
    case "SRK_MASTER_LIST_DATA":
      handleMasterListData(event.data.data);
      break;
  }
});
```

### 2. Respond to Extension Check

```javascript
function handleExtensionCheck() {
  console.log("Extension check received from add-in");

  // Respond that extension is installed
  window.postMessage({
    type: "SRK_EXTENSION_INSTALLED"
  }, "*");
}
```

### 3. Respond to Master List Request

```javascript
function handleMasterListRequest(message) {
  console.log("Master List request received:", message.timestamp);

  // Decide if you want the data
  const needsUpdate = checkIfDataNeeded(); // Your logic here

  // Send response
  window.postMessage({
    type: "SRK_MASTER_LIST_RESPONSE",
    wantsData: needsUpdate
  }, "*");
}

function checkIfDataNeeded() {
  // Example: Always accept data on first load
  const hasExistingData = localStorage.getItem('masterListData');
  return !hasExistingData || shouldRefresh();
}
```

### 4. Process Master List Data

```javascript
function handleMasterListData(data) {
  console.log(`Received Master List with ${data.totalStudents} students`);
  console.log(`Available columns: ${data.headers.join(', ')}`);
  console.log(`Data timestamp: ${data.timestamp}`);

  // Store the data
  localStorage.setItem('masterListData', JSON.stringify(data));
  localStorage.setItem('lastUpdated', data.timestamp);

  // Process students - access columns dynamically using header names
  data.students.forEach(student => {
    console.log(`Student: ${student["Student Name"]}`);

    // Access gradebook URL if available (URL extracted from HYPERLINK formula)
    if (student.Gradebook) {
      console.log(`  Gradebook: ${student.Gradebook}`);
    }

    // Check if student needs attention
    if (student["Days Out"] && student["Days Out"] > 5) {
      console.log(`  ⚠️ High absence: ${student["Days Out"]} days`);
    }

    // Access any custom columns
    if (student.Gender) {
      console.log(`  Gender: ${student.Gender}`);
    }
    if (student.ProgramVersion) {
      console.log(`  Program: ${student.ProgramVersion}`);
    }
  });

  // Update your UI or trigger your automation
  updateExtensionUI(data);
}
```

### 5. Example: Find Student by Name

```javascript
function findStudent(name) {
  const data = JSON.parse(localStorage.getItem('masterListData'));
  if (!data) return null;

  const normalizedSearch = name.toLowerCase();
  return data.students.find(student => {
    const studentName = student["Student Name"] || "";
    return studentName.toLowerCase().includes(normalizedSearch);
  });
}

// Usage
const student = findStudent("John Smith");
if (student && student.Gradebook) {
  // Open gradebook in new tab
  chrome.tabs.create({ url: student.Gradebook });
}

// Access any column dynamically
if (student) {
  console.log(`Student: ${student["Student Name"]}`);
  console.log(`ID: ${student.SyStudentId}`);
  console.log(`Grade: ${student.Grade}`);
  console.log(`Gender: ${student.Gender}`);
  console.log(`Shift: ${student.Shift}`);
}
```

### 6. Importing Master List Data to Excel

The Chrome extension can send master list data to import into the add-in's Excel sheet:

```javascript
function importMasterListToExcel(students) {
  // Prepare headers (column names)
  const headers = [
    "StudentName",
    "Student ID",
    "Student Number",
    "Grade",
    "Phone",
    "Email",
    "Days Out"
  ];

  // Prepare data rows (must match header order)
  const data = students.map(student => [
    student.name,           // StudentName
    student.id,             // Student ID
    student.number,         // Student Number
    student.grade,          // Grade
    student.phone,          // Phone
    student.email,          // Email
    student.daysOut         // Days Out
  ]);

  // Send import message to add-in
  window.postMessage({
    type: "SRK_IMPORT_MASTER_LIST",
    data: {
      headers: headers,
      data: data
    }
  }, "*");

  console.log(`Sent ${data.length} students to Excel for import`);
}

// Example usage with sample data
const studentsToImport = [
  {
    name: "Smith, John",
    id: "12345",
    number: "STU001",
    grade: 85.5,
    phone: "555-1234",
    email: "john@school.edu",
    daysOut: 3
  },
  {
    name: "Doe, Jane",
    id: "67890",
    number: "STU002",
    grade: 72,
    phone: "555-9876",
    email: "jane@school.edu",
    daysOut: 7
  }
];

importMasterListToExcel(studentsToImport);
```

**Important Notes for Import:**
- The add-in will only import if a "Master List" sheet already exists
- Column headers are matched case-insensitively with whitespace normalization (e.g., "StudentName" matches "Student Name")
- Gradebook URLs (starting with http:// or https://) are automatically wrapped in HYPERLINK formulas with "Grade Book" as the display text
- Grade column receives automatic 3-color conditional formatting (Red for low, Yellow for medium ~70%, Green for high)
- The formatting automatically detects 0-1 or 0-100 grade scales
- Existing Gradebook links and Assigned values are preserved for existing students
- New students will be highlighted in light blue
- All data must be in array format matching the headers order

## Important Notes

1. **All Columns Included**: The add-in now sends ALL columns from the Master List sheet, not just predefined ones. This means any custom columns you add will automatically be included in the payload.

2. **Column Mapping**: The `columnMapping` object maps each header name to its column index (0-based). Use this if you need to reference data back to specific columns. The `headers` array provides all column names in order.

3. **Dynamic Column Access**: Student objects use the actual header names as keys. Use bracket notation to access columns with spaces (e.g., `student["Student Name"]`) or dot notation for columns without spaces (e.g., `student.Grade`).

4. **Optional Fields**: All fields in student objects are optional (except the student name which is required). Always check if a field exists before using it.

5. **Gradebook URLs**: The `Gradebook` field contains the full URL extracted from Excel HYPERLINK formulas, making it easy to open student gradebooks directly. The formula is automatically parsed to extract just the URL.

6. **Timestamps**: Use the `timestamp` field to track when data was last synced and decide if you need to request fresh data.

7. **Message Origin**: Always validate message origins in production for security:
   ```javascript
   window.addEventListener("message", (event) => {
     // Validate origin if needed
     // if (event.origin !== "https://expected-origin.com") return;

     // Process message
   });
   ```

## Message Types Reference

| Message Type | Direction | Purpose |
|-------------|-----------|---------|
| `SRK_CHECK_EXTENSION` | Add-in → Extension | Ping to detect extension |
| `SRK_EXTENSION_INSTALLED` | Extension → Add-in | Confirm extension is present |
| `SRK_REQUEST_MASTER_LIST` | Add-in → Extension | Ask if extension wants data |
| `SRK_MASTER_LIST_RESPONSE` | Extension → Add-in | Accept or decline data |
| `SRK_MASTER_LIST_DATA` | Add-in → Extension | Send Master List payload |
| `SRK_IMPORT_MASTER_LIST` | Extension → Add-in | Import master list data into Excel |

## Testing

To test the integration:

1. Open Excel with the add-in loaded
2. Open browser DevTools Console
3. Load your Chrome extension
4. Watch for messages in console:
   - "Background script initialized - starting Chrome Extension Service"
   - "Background: Chrome extension detected and ready!"
   - "Background: Asking extension if it wants Master List data..."
   - "TransferMasterList: Successfully read X students from Master List"

5. In your extension, verify you receive the data payload
