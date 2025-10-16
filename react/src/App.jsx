// Timestamp: 2025-10-02 04:22 PM | Version: 1.2.0
import React from 'react';
// The import path is updated to look inside the new folder.
import StudentView from './components/studentView/StudentView.jsx';
import About from './components/about/About.jsx';
import { ToastContainer } from 'react-toastify';
import 'react-toastify/dist/ReactToastify.css';
import { Slide, Zoom, Flip, Bounce } from 'react-toastify';

function App() {
  return (
    <>
      <StudentView />
      <ToastContainer
      transition={Slide}
       />
    </>
  );
}

export default App;

