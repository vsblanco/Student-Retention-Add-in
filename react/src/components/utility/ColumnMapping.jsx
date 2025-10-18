export const Sheets = {
  HISTORY: "Student History",
  MISSING_ASSIGNMENT: "Missing Assignments"
};

export const COLUMN_ALIASES = {
  StudentName: ['Student Name', 'Student'],
  ID: ['Student ID', 'Student Number','Student identifier'],
  Gender: ['Gender'],
  Phone: ['Phone Number', 'Contact'],
  CreatedBy: ['Created By', 'Author', 'Advisor'],
  OtherPhone: ['Other Phone', 'Alt Phone'],
  StudentEmail: ['Email', 'Student Email'],
  PersonalEmail: ['Other Email'],
  Assigned: ['Advisor'],
  ExpectedStartDate: ['Expected Start Date', 'Start Date','ExpStartDate'],
  Grade: ['Current Grade', 'Grade %', 'Grade'],
  LDA: ['Last Date of Attendance', 'LDA'],
  DaysOut: ['Days Out'],
  Gradebook: ['Gradebook','gradeBookLink'],
  MissingAssignments: ['Missing Assignments', 'Missing'],
  Outreach: ['Outreach', 'Comments', 'Notes', 'Comment']
  // You can add more aliases for other columns here
};

export const COLUMN_ALIASES_ASSIGNMENTS = {
  StudentName: ['Student Name', 'Student'],
  title: ['Assignment Title', 'Title', 'Assignment'],
  dueDate: ['Due Date', 'Deadline'],
  score: ['Score', 'Points'],
  submissionLink: ['Submission Link', 'Submission', 'Submit Link'],
  assignmentLink: ['Assignment Link', 'Assignment URL', 'Assignment Page', 'Link'],
  Gradebook: ['Gradebook','gradeBookLink'],
};

export const COLUMN_ALIASES_HISTORY = {
  timestamp: ['Timestamp', 'Date', 'Time', 'Created At'],
  comment: ['Comment', 'Notes', 'History', 'Entry'],
  createdBy: ['Created By', 'Author', 'Advisor'],
  tag: ['Tag', 'Category', 'Type','Tags'],
  StudentID: ['Student ID', 'Student Number','Student identifier']
  // Add more aliases as needed for history columns
};
