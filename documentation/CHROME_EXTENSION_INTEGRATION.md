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
  - Match incoming headers to Master List headers (case-insensitive)
  - Preserve existing Gradebook hyperlinks for students already in the list
  - Preserve "Assigned" column values and their colors
  - Highlight new students in light blue (#ADD8E6)
  - Clear and repopulate the sheet with the imported data

## Payload Structure

### Master List Data Payload (Add-in → Extension)

```typescript
interface MasterListPayload {
  type: "SRK_MASTER_LIST_DATA";
  data: {
    sheetName: string;              // Always "Master List"
    columnMapping: ColumnMapping;   // Maps fields to column indices
    students: Student[];            // Array of student records
    totalStudents: number;          // Total count
    timestamp: string;              // ISO 8601 timestamp
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

interface ColumnMapping {
  studentName: number;      // Column index for Student Name
  syStudentId: number;      // Column index for SyStudentId (School SIS ID)
  studentNumber: number;    // Column index for Student Number (School ID)
  gradeBook: number;        // Column index for Gradebook link
  daysOut: number;          // Column index for Days Out
  lastLda: number;          // Column index for Last LDA
  grade: number;            // Column index for Grade/Score
  primaryPhone: number;     // Column index for Primary Phone
  otherPhone: number;       // Column index for Other Phone
  personalEmail: number;    // Column index for Personal Email
  studentEmail: number;     // Column index for Student Email
  assigned: number;         // Column index for Assigned advisor
  outreach: number;         // Column index for Outreach notes
}

interface Student {
  studentName: string;           // Required - student's name
  syStudentId?: string;          // School's SIS (Student Information System) ID
  studentNumber?: string;        // School's student identification number
  gradeBook?: string;            // URL to gradebook (extracted from HYPERLINK formula)
  daysOut?: number;              // Days since last attendance
  lastLda?: any;                 // Last LDA date/value
  grade?: number | string;       // Current grade/score
  primaryPhone?: string;         // Primary contact phone
  otherPhone?: string;           // Secondary phone
  personalEmail?: string;        // Personal email address
  studentEmail?: string;         // School email address
  assigned?: string;             // Assigned advisor/counselor
  outreach?: string;             // Outreach notes
}
```

## Example Payload

```json
{
  "type": "SRK_MASTER_LIST_DATA",
  "data": {
    "sheetName": "Master List",
    "columnMapping": {
      "studentName": 0,
      "syStudentId": 1,
      "studentNumber": 2,
      "gradeBook": 5,
      "daysOut": 8,
      "lastLda": 9,
      "grade": 4,
      "primaryPhone": 10,
      "otherPhone": 11,
      "personalEmail": 12,
      "studentEmail": 3,
      "assigned": 6,
      "outreach": 7
    },
    "students": [
      {
        "studentName": "Smith, John",
        "syStudentId": "12345",
        "studentNumber": "STU001",
        "studentEmail": "john.smith@school.edu",
        "grade": 85.5,
        "gradeBook": "https://canvas.instructure.com/courses/123/gradebook/speed_grader?assignment_id=456&student_id=12345",
        "assigned": "Dr. Johnson",
        "outreach": "Called on 12/15",
        "daysOut": 3,
        "lastLda": "12/10/2025",
        "primaryPhone": "555-1234",
        "otherPhone": "555-5678",
        "personalEmail": "john@email.com"
      },
      {
        "studentName": "Doe, Jane",
        "syStudentId": "67890",
        "studentNumber": "STU002",
        "studentEmail": "jane.doe@school.edu",
        "grade": 72,
        "gradeBook": "https://canvas.instructure.com/courses/123/gradebook/speed_grader?assignment_id=456&student_id=67890",
        "assigned": "Dr. Williams",
        "daysOut": 7,
        "lastLda": "12/05/2025",
        "primaryPhone": "555-9876"
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
  console.log(`Data timestamp: ${data.timestamp}`);

  // Store the data
  localStorage.setItem('masterListData', JSON.stringify(data));
  localStorage.setItem('lastUpdated', data.timestamp);

  // Process students
  data.students.forEach(student => {
    console.log(`Student: ${student.studentName}`);

    // Access gradebook URL if available
    if (student.gradeBook) {
      console.log(`  Gradebook: ${student.gradeBook}`);
    }

    // Check if student needs attention
    if (student.daysOut && student.daysOut > 5) {
      console.log(`  ⚠️ High absence: ${student.daysOut} days`);
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
  return data.students.find(student =>
    student.studentName.toLowerCase().includes(normalizedSearch)
  );
}

// Usage
const student = findStudent("John Smith");
if (student && student.gradeBook) {
  // Open gradebook in new tab
  chrome.tabs.create({ url: student.gradeBook });
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
- Column headers are matched case-insensitively
- Existing Gradebook links and Assigned values are preserved for existing students
- New students will be highlighted in light blue
- All data must be in array format matching the headers order

## Important Notes

1. **Column Mapping**: The `columnMapping` object tells you which column index each field came from. This is useful if you need to map data back to the spreadsheet.

2. **Optional Fields**: Most fields in the `Student` object are optional. Always check if a field exists before using it.

3. **Gradebook URLs**: The `gradeBook` field contains the full URL extracted from Excel HYPERLINK formulas, making it easy to open student gradebooks directly.

4. **Timestamps**: Use the `timestamp` field to track when data was last synced and decide if you need to request fresh data.

5. **Message Origin**: Always validate message origins in production for security:
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
