// Timestamp: 2025-10-02 04:18 PM | Version: 1.1.0
import React from 'react';
// CSS import is removed.

function StudentDetails() {
  // This component will eventually receive student data via props.

  // --- Inline Styles ---
  const headingStyles = {
    marginTop: 0,
    color: '#2c3e50',
  };

  return (
    <div className="student-details">
      <h2 style={headingStyles}>Student Details</h2>
      <p>Details about the selected student will go here.</p>
    </div>
  );
}

export default StudentDetails;

