// Timestamp: 2025-10-02 04:22 PM | Version: 1.2.0
import React from 'react';
// The import path is updated to look inside the new folder.
import StudentView from './components/studentView/StudentView.jsx';
import About from './components/about/About.jsx';

function App() {
  return (
    <div className="App flex flex-row">
      <StudentView />
    </div>
  );
}

export default App;

