// Timestamp: 2025-10-02 03:42 PM | Version: 1.0.0
import React from 'react'
import ReactDOM from 'react-dom/client'
import App from './App.jsx'
import './index.css'

/*
 * This is the key change. We are telling our React app to wait
 * until the Office host (Excel) is ready before we try to render anything.
 */
Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    ReactDOM.createRoot(document.getElementById('root')).render(
      <React.StrictMode>
        <App />
      </React.StrictMode>
    );
  }
});
