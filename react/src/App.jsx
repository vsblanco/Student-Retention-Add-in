// Timestamp: 2025-12-11 12:15:00 | Version: 1.4.0
import React, { useState, useEffect, useRef } from 'react';
import StudentView from './components/studentView/StudentView.jsx';
import About from './components/about/About.jsx';
import { ToastContainer } from 'react-toastify';
import 'react-toastify/dist/ReactToastify.css';
import { Slide } from 'react-toastify';
import Settings from './components/settings/Settings.jsx';
import ImportManager from './components/importData/ImportManager.jsx';
import LDAManager from './components/createLDA/LDAManager.jsx';

function App() {
  const [page, setPage] = useState(null);
  const [isLoading, setIsLoading] = useState(true);

  // We use a ref to prevent double-firing the ready state
  const isReadyRef = useRef(false);

  // The signal function that child components can call when they are fully loaded
  const handleFeatureReady = () => {
    if (!isReadyRef.current) {
      isReadyRef.current = true;
      setIsLoading(false);
    }
  };

  useEffect(() => {
    const params = new URLSearchParams(window.location.search);
    const pageParam = params.get('page');
    const targetPage = pageParam || 'student-view';
    
    setPage(targetPage);

    // LOGIC: Check if the target page provides its own "Ready" signal.
    // Only 'student-view' implements the handshake. 
    // All other pages continue as normal with a default timeout.
    if (targetPage !== 'student-view') {
      const timer = setTimeout(() => {
        handleFeatureReady();
      }, 600); // Default fade out for non-async pages
      return () => clearTimeout(timer);
    }
    
    // Safety Fallback: If StudentView crashes or takes too long (> 8 seconds),
    // force remove the overlay so the user isn't stuck forever.
    const safetyTimer = setTimeout(() => {
      if (!isReadyRef.current) {
        console.warn("Feature took too long to report ready. Forcing load.");
        handleFeatureReady();
      }
    }, 8000);

    return () => clearTimeout(safetyTimer);
  }, []);

  // Keep-Alive Heartbeat (Unchanged)
  useEffect(() => {
    const keepAliveInterval = setInterval(() => {
      if (typeof chrome !== 'undefined' && chrome.runtime && chrome.runtime.sendMessage) {
        chrome.runtime.sendMessage({ type: 'keep_alive_ping' }, (response) => {
          if (chrome.runtime.lastError) { /* ignore */ }
        });
      }
    }, 20000); 
    return () => clearInterval(keepAliveInterval);
  }, []);

  const renderPage = () => {
    if (page === null) return null;
    switch (page) {
      case 'settings':
        return <Settings />;
      case 'import':
        return <ImportManager />;
      case 'about':
        return <About />;
      case 'create-lda':
        return <LDAManager />;
      case 'student-view':
        // Pass the handshake prop ONLY to StudentView
        return <StudentView onReady={handleFeatureReady} />;
      default:
        return <StudentView onReady={handleFeatureReady} />;
    }
  };

  return (
    <>
      {/* LOADING OVERLAY */}
      <div 
        style={{
          position: 'fixed',
          inset: 0,
          zIndex: 9999,
          backgroundColor: '#f9fafb',
          display: 'flex',
          alignItems: 'center',
          justifyContent: 'center',
          transition: 'opacity 0.7s ease-out',
          opacity: isLoading ? 1 : 0,
          pointerEvents: isLoading ? 'auto' : 'none',
        }}
      >
        <div className="flex flex-col items-center gap-4">
           <div className="animate-spin rounded-full h-12 w-12 border-t-4 border-b-4 border-blue-600"></div>
           <span className="text-gray-500 font-medium text-sm animate-pulse">Loading S.H.A.R.K...</span>
        </div>
      </div>

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