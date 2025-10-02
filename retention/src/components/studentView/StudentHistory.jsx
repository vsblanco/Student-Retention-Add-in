// Timestamp: 2025-10-02 04:37 PM | Version: 3.0.0
import React from 'react';

function StudentHistory({ history }) {

  // This function parses the raw history string from the Excel cell
  // into an array of objects, just like the original taskpane.js.
  const parseHistory = (historyString) => {
    if (!historyString || typeof historyString !== 'string') {
      return [];
    }
    const entries = [];
    const lines = historyString.split('\n').filter(line => line.trim() !== '');
    for (let i = 0; i < lines.length; i += 2) {
      if (lines[i] && lines[i+1]) {
        entries.push({
          date: lines[i].trim(),
          comment: lines[i+1].trim()
        });
      }
    }
    return entries;
  };

  const historyEntries = parseHistory(history);

  // --- STYLES ---
  const noHistoryStyles = {
    color: '#7f8c8d',
    textAlign: 'center',
    marginTop: '20px'
  };

  const historyItemStyles = {
    border: '1px solid #ecf0f1',
    borderRadius: '4px',
    padding: '10px',
    marginBottom: '10px',
    backgroundColor: '#f9f9f9'
  };
  
  const dateStyles = {
    fontWeight: 'bold',
    color: '#34495e',
    fontSize: '13px',
    marginBottom: '5px'
  };

  const commentStyles = {
    color: '#7f8c8d',
    fontSize: '14px',
    whiteSpace: 'pre-wrap', // Preserves formatting from the cell
    margin: 0
  };

  if (historyEntries.length === 0) {
    return <p style={noHistoryStyles}>No history found for this student.</p>;
  }

  return (
    <div>
      {historyEntries.map((entry, index) => (
        <div key={index} style={historyItemStyles}>
          <p style={dateStyles}>{entry.date}</p>
          <p style={commentStyles}>{entry.comment}</p>
        </div>
      ))}
    </div>
  );
}

export default StudentHistory;

