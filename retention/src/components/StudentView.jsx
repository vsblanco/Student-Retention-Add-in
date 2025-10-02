// Timestamp: 2025-10-02 04:15 PM | Version: 1.1.0
import React from 'react';
import StudentDetails from './StudentDetails.jsx';
import StudentHistory from './StudentHistory.jsx';
// We no longer need to import the separate CSS file.

function StudentView() {
  // --- Inline Styles ---
  // We define our styles as JavaScript objects.
  // CSS properties like 'padding' are written as is.
  const containerStyles = {
    padding: '1rem',
  };

  // CSS properties with a dash, like 'border-top',
  // become camelCase, like 'borderTop'.
  const dividerStyles = {
    margin: '1.5rem 0',
    border: 'none',
    borderTop: '1px solid #eee',
  };

  return (
    // We apply the style objects directly to the elements.
    <div style={containerStyles} className="student-view-container">
      <StudentDetails />
      <hr style={dividerStyles} className="divider" />
      <StudentHistory />
    </div>
  );
}

export default StudentView;

