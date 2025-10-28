// Timestamp: 2025-10-02 03:42 PM | Version: 1.0.0
import React from 'react'
import ReactDOM from 'react-dom/client'
import App from './App.jsx'
import Taskpane from './components/utility/taskpane.jsx'
import Ribbon from './components/utility/Ribbon.jsx'
import './index.css'

/*
 * This is the key change. We are telling our React app to wait
 * until the Office host (Excel) is ready before we try to render anything.
 */

function renderApp(content) {
  const rootEl = document.getElementById('root');
  if (rootEl) {
    ReactDOM.createRoot(rootEl).render(content);
  } else {
    document.body.innerHTML = '<h1 style="color:red">No #root element found</h1>';
  }
}

if (typeof window.Office === 'undefined') {
  // Not running inside Office, render app for browser testing
  renderApp(
    <React.StrictMode>
      <App />
    </React.StrictMode>
  );
} else {
  Office.onReady((info) => {
    if (info.host === Office.HostType.Excel) {
      renderApp(
        <React.StrictMode>
          <App />
        </React.StrictMode>
      );
    } else {
      renderApp(
        <>
          <Ribbon />
          <Taskpane
            open={true}
            onClose={() => {}}
            header={<div>Student Retention Add-in</div>}
          >
            <App />
          </Taskpane>
        </>
      );
    }
  });
}

// Minimal entry module to render a simple app into #root
const root = document.getElementById('root');
if (!root) {
  console.error('Root element not found');
} else {
  root.innerHTML = `
    <div style="font-family: Segoe UI, Roboto, Arial, sans-serif; padding: 16px;">
      <h1>Student Retention Settings</h1>
      <p>The app script loaded successfully from ./src/main.jsx</p>
    </div>
  `;
}