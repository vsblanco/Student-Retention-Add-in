// Timestamp: 2025-10-02 03:42 PM | Version: 1.0.0
import React from 'react'
import ReactDOM from 'react-dom/client'
import App from './App.jsx'
import StudentView from './components/studentView/StudentView.jsx'
import './index.css'

/*
 * This is the key change. We are telling our React app to wait
 * until the Office host (Excel) is ready before we try to render anything.
 */

if (typeof window.Office === 'undefined') {
  // Not running inside Office, render app for browser testing
  const rootEl = document.getElementById('root');
  if (rootEl) {
    ReactDOM.createRoot(rootEl).render(
      <React.StrictMode>
        <App />
      </React.StrictMode>
    );
  } else {
    document.body.innerHTML = '<h1 style="color:red">No #root element found</h1>';
  }
} else {
  Office.onReady((info) => {
    const rootEl = document.getElementById('root');
    if (info.host === Office.HostType.Excel) {
      if (rootEl) {
        ReactDOM.createRoot(rootEl).render(
          <React.StrictMode>
            <App />
          </React.StrictMode>
        );
      }
    } else {
      // Office.js loaded, but not inside Office client
      if (rootEl) {
        ReactDOM.createRoot(rootEl).render(
          <>
            <StudentView />
          </>
        );
      } else {
        document.body.innerHTML = '<h1 style="color:red">No #root element found</h1>';
      }
    }
  });
}
