// Timestamp: 2025-10-02 04:22 PM | Version: 1.2.0
import React from 'react';
// The import path is updated to look inside the new folder.
import StudentView from './components/studentView/StudentView.jsx';
import About from './components/about/About.jsx';
import { ToastContainer } from 'react-toastify';
import 'react-toastify/dist/ReactToastify.css';
import { Slide, Zoom, Flip, Bounce } from 'react-toastify';
import Settings from './components/settings/Settings.jsx';

function App() {
  return (
    <>
      <Settings />
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

