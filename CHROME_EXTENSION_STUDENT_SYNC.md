# Chrome Extension Student Selection Sync

## Overview

The Student Retention Add-in now sends real-time student selection updates to the Chrome extension. When a user selects one or more students in Excel, the add-in automatically sends the selected student data to the extension via `window.postMessage()`.

## Message Type

**Message Type:** `SRK_SELECTED_STUDENTS`

This message is sent whenever:
- A single student row is selected
- Multiple student rows are selected
- The selection changes to a different student/students

## Payload Structure

### Top-Level Message

```javascript
{
  type: "SRK_SELECTED_STUDENTS",
  data: {
    students: [...],      // Array of student objects
    count: 1,             // Number of students selected
    timestamp: "2025-12-23T10:30:01.500Z"  // ISO 8601 timestamp
  }
}
```

### Student Object Structure

Each student in the `students` array contains the following fields:

```javascript
{
  name: "Jane Doe",               // Student's full name
  syStudentId: "123456",          // SyStudentID / Canvas ID
  phone: "555-1234",              // Primary phone number
  otherPhone: "555-5678"          // Secondary/alternate phone number
}
```

## Field Details

| Field | Type | Description | Source Column |
|-------|------|-------------|---------------|
| `name` | String | Student's full name | `StudentName`, `Student Name`, or `Student` |
| `syStudentId` | String | System/Canvas student ID | `ID`, `Student ID`, `SyStudentID`, or `Student identifier` |
| `phone` | String | Primary contact phone | `Phone`, `Phone Number`, or `Contact` |
| `otherPhone` | String | Alternate contact phone | `OtherPhone`, `Other Phone`, or `Alt Phone` |

**Note:** All fields will be empty strings (`""`) if the data is not available in the Excel sheet.

## Example Payloads

### Single Student Selection

```javascript
{
  type: "SRK_SELECTED_STUDENTS",
  data: {
    students: [
      {
        name: "John Smith",
        syStudentId: "987654",
        phone: "555-0123",
        otherPhone: "555-9876"
      }
    ],
    count: 1,
    timestamp: "2025-12-23T14:25:30.120Z"
  }
}
```

### Multiple Student Selection

```javascript
{
  type: "SRK_SELECTED_STUDENTS",
  data: {
    students: [
      {
        name: "Alice Johnson",
        syStudentId: "111222",
        phone: "555-1111",
        otherPhone: "555-2222"
      },
      {
        name: "Bob Williams",
        syStudentId: "333444",
        phone: "555-3333",
        otherPhone: ""
      },
      {
        name: "Carol Martinez",
        syStudentId: "555666",
        phone: "555-5555",
        otherPhone: "555-6666"
      }
    ],
    count: 3,
    timestamp: "2025-12-23T14:27:45.890Z"
  }
}
```

## Chrome Extension Implementation

### Listening for Messages

To receive these messages in your Chrome extension, add a message listener:

```javascript
// In your content script or background script
window.addEventListener("message", (event) => {
  // Validate the message
  if (!event.data || !event.data.type) return;

  // Handle student selection sync
  if (event.data.type === "SRK_SELECTED_STUDENTS") {
    const { students, count, timestamp } = event.data.data;

    console.log(`Received ${count} student(s) at ${timestamp}`);

    // Process the student data
    students.forEach(student => {
      console.log(`Student: ${student.name} (ID: ${student.syStudentId})`);
      console.log(`  Primary Phone: ${student.phone}`);
      console.log(`  Other Phone: ${student.otherPhone}`);
    });

    // Your implementation here
    syncStudentsToExtension(students);
  }
});
```

### Example: Storing in Chrome Storage

```javascript
function syncStudentsToExtension(students) {
  // Store in chrome.storage.local
  chrome.storage.local.set({
    selectedStudents: students,
    lastSync: new Date().toISOString()
  }, () => {
    console.log('Students synced to extension storage');
  });

  // Or send to background script
  chrome.runtime.sendMessage({
    action: 'updateSelectedStudents',
    students: students
  });
}
```

### Example: Updating Extension UI

```javascript
function updateExtensionUI(students) {
  const studentList = document.getElementById('student-list');

  if (students.length === 0) {
    studentList.innerHTML = '<p>No students selected</p>';
    return;
  }

  const listHTML = students.map(student => `
    <div class="student-card">
      <h3>${student.name}</h3>
      <p>ID: ${student.syStudentId}</p>
      <p>Phone: ${student.phone || 'N/A'}</p>
      ${student.otherPhone ? `<p>Alt Phone: ${student.otherPhone}</p>` : ''}
    </div>
  `).join('');

  studentList.innerHTML = listHTML;
}
```

## Message Flow Diagram

```
Excel Selection Changed
        ↓
StudentView.jsx (React)
        ↓
chromeExtensionService.sendSelectedStudents()
        ↓
window.postMessage(SRK_SELECTED_STUDENTS)
        ↓
Chrome Extension (Content Script)
        ↓
Extension UI Update / Storage / Background Script
```

## Important Notes

1. **Real-time Updates:** Messages are sent immediately when selection changes (debounced at 250ms to prevent lag)

2. **Empty Fields:** If a student doesn't have a phone number or other field, it will be an empty string (`""`), not `null` or `undefined`

3. **No Selection:** When the header row is selected or no valid student is selected, no message is sent

4. **Message Origin:** Always validate the message origin in production to ensure security

5. **Existing Communication:** This is an addition to the existing `SRK_CHECK_EXTENSION` and `SRK_MASTER_LIST_DATA` message types

6. **Field Mapping:** The add-in uses flexible column name matching (case-insensitive), so these fields should always be populated if the data exists in the Excel sheet

## Testing

To test the integration:

1. Open the Student Retention Add-in in Excel
2. Install your Chrome extension with the message listener
3. Select a student row in Excel
4. Check your extension's console for the `SRK_SELECTED_STUDENTS` message
5. Select multiple rows to test batch updates
6. Verify the student data is correctly parsed and displayed

## Security Considerations

- **Data Validation:** Always validate the incoming data structure before using it
- **Origin Validation:** In production, validate the message origin to prevent XSS attacks
- **PII Handling:** Phone numbers and student names are personally identifiable information (PII) - handle with appropriate security measures
- **Storage:** If storing student data, ensure compliance with privacy policies and data retention requirements

## Version History

- **v1.0.0** (2025-12-23): Initial implementation of student selection sync
  - Single and multiple student selection support
  - Four core fields: name, syStudentId, phone, otherPhone
  - Real-time sync with debouncing

## Support

For questions or issues with this integration, please refer to the main Student Retention Add-in documentation or contact the development team.
