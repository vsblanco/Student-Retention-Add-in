// Timestamp: 2025-10-02 04:45 PM | Version: 1.2.1
import React, { useState, useEffect } from 'react';
// The import path is updated to look inside the new folder.
import StudentView from './components/studentView/StudentView.jsx';
import About from './components/about/About.jsx';
import { ToastContainer } from 'react-toastify';
import 'react-toastify/dist/ReactToastify.css';
import { Slide, Zoom, Flip, Bounce } from 'react-toastify';
import Settings from './components/settings/Settings.jsx';
import ImportManager from './components/importData/ImportManager.jsx';
import LDAManager from './components/createLDA/LDAManager.jsx';

function App() {
  // start as null so no page mounts until we parse the URL
  const [page, setPage] = useState(null);

  useEffect(() => {
    // Parse the query parameter ?page=... from the URL
    const params = new URLSearchParams(window.location.search);
    const pageParam = params.get('page');
    // set parsed page or default to 'student-view'
    setPage(pageParam || 'student-view');
  }, []);

  // Determine which component to render based on the current page state
  const renderPage = () => {
    // while null, don't mount any page (prevents StudentView background work)
    if (page === null) return null;
    switch (page) {
      case 'settings':
        return <Settings />;
      case 'import':
        return <ImportManager />;
      case 'about':
        return <About />;
      case 'student-view':
        return <StudentView />;
      case 'create-lda':
        return <LDAManager />;
      default:
        return <StudentView />;
    }
  };

  return (
    <>
      {renderPage()}
      <ToastContainer
        position="bottom-left"
        pauseOnFocusLoss={false}
        transition={Slide}
        limit={3}
      />
    </>
  );
}

export default App;