const ExampleStudent = {
  StudentName: "Jane Doe",
  ID: "123456",
  Gender: "Boy",
  Phone: "555-1234",
  OtherPhone: "555-5678",
  StudentEmail: "jane.doe@university.edu",
  PersonalEmail: "jane.doe@gmail.com",
  Assigned: "Dr. Smith",
  DaysOut: 10,
  Grade: "77%",
  LDA: "2024-05-20",
  Gradebook: "https://google.com",
  // History as an array of objects for test mode
  History: [
    {
      timestamp: "2024-01-15",
      comment: "Advised: Discussed course selection.",
      ID: "123456",
      studentName: "Jane Doe",
      createdBy: "Dr. Smith",
      tag: "Outreach"
    },
    {
      timestamp: "2024-01-15",
      comment: "Left Voicemail",
      ID: "123456",
      studentName: "Jane Doe",
      createdBy: "Dr. Smith",
    },
    {
      timestamp: "2024-03-10",
      comment: "Follow-up: Checked on progress.",
      ID: "123456",
      studentName: "Jane Doe",
      createdBy: "Dr. Smith",
      tag: "Contacted"
    }
  ]
  // Optionally add an alias for testing alias logic:
  // Notes: [...]
};

export default ExampleStudent;
