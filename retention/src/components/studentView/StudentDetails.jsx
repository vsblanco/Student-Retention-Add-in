// Timestamp: 2025-10-02 04:37 PM | Version: 3.0.0
import React from 'react';

// A small reusable component for displaying a single detail item.
// This keeps our main component clean.
const DetailItem = ({ label, value }) => {
  const itemStyles = {
    marginBottom: '10px'
  };
  const labelStyles = {
    fontWeight: '600',
    color: '#34495e',
    display: 'block',
    fontSize: '14px'
  };
  const valueStyles = {
    color: '#7f8c8d',
    fontSize: '14px',
    marginLeft: '5px'
  };

  return (
    <div style={itemStyles}>
      <span style={labelStyles}>{label}:</span>
      <span style={valueStyles}>{value || 'N/A'}</span>
    </div>
  );
};

function StudentDetails({ student }) {
  // --- STYLES ---
  const studentNameStyles = {
    fontSize: '22px',
    fontWeight: '600',
    color: '#2c3e50',
    marginBottom: '15px',
    borderBottom: '1px solid #ecf0f1',
    paddingBottom: '10px'
  };

  const sectionStyles = {
    marginBottom: '20px'
  };

  // Check for both name variations from the student object
  const studentName = student["StudentName"] || student["Student Name"];

  return (
    <div>
      <h2 style={studentNameStyles}>{studentName}</h2>
      
      <div style={sectionStyles}>
        <DetailItem label="Student ID" value={student.ID} />
        <DetailItem label="Primary Phone" value={student.Phone} />
        <DetailItem label="Other Phone" value={student.OtherPhone} />
        <DetailItem label="Assigned" value={student.Assigned} />
        <DetailItem label="Student Email" value={student.StudentEmail} />
        <DetailItem label="Personal Email" value={student.PersonalEmail} />
      </div>

      <div style={sectionStyles}>
        <DetailItem label="Last LDA" value={student["LDA"]} />
      </div>

      {/* Placeholder for future actions */}
      {/* <div id="actions-container">
        <h3>Actions</h3>
      </div> */}
    </div>
  );
}

export default StudentDetails;

