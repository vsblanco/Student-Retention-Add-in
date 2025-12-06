// Timestamp: 2025-12-03 12:45:00 | Version: 1.2.2
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

  // ---------------------------------------------------------------------------
  // GLOBAL FIX: Keep-Alive Heartbeat
  // ---------------------------------------------------------------------------
  // This effect runs once on mount. It pings the runtime environment every 20s
  // to prevent the background script/service worker from going to sleep.
  useEffect(() => {
    const keepAliveInterval = setInterval(() => {
      // Check if the chrome runtime API is available
      if (typeof chrome !== 'undefined' && chrome.runtime && chrome.runtime.sendMessage) {
        
        // Send a lightweight 'ping' message
        chrome.runtime.sendMessage({ type: 'keep_alive_ping' }, (response) => {
          
          // CRITICAL: We must check chrome.runtime.lastError in the callback.
          // Accessing this property marks the error as "checked" and prevents the
          // "Unchecked runtime.lastError: Could not establish connection" noise in the console.
          if (chrome.runtime.lastError) {
            // Context is likely invalid or background is dead. 
            // We swallow the error here so it doesn't crash the UI logic.
            // console.warn("Background connection lost (handled):", chrome.runtime.lastError.message);
          }
        });
      }
    }, 20000); // 20 seconds is usually frequent enough to prevent sleep

    // Cleanup interval on unmount
    return () => clearInterval(keepAliveInterval);
  }, []);
  // ---------------------------------------------------------------------------

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