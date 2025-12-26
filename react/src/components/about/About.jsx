// Timestamp: 2025-11-22 | Version: 3.0.0
import React, { useState, useEffect } from "react";
import MarkdownViewer from "../utility/MarkdownViewer";

const About = ({ onReady } = {}) => {
  const BASE_URL = import.meta.env.BASE_URL;
  const HOME_FILE = "about.md";

  // State to track which markdown file is currently displayed
  const [currentFileName, setCurrentFileName] = useState(HOME_FILE);

  // Construct the full path for fetching
  const fileUrl = `${BASE_URL}${currentFileName}`;

  const handleLinkClick = (href) => {
    // Remove any leading slash if present to keep filenames clean
    const fileName = href.startsWith("/") ? href.slice(1) : href;
    setCurrentFileName(fileName);
  };

  const goHome = () => setCurrentFileName(HOME_FILE);

  // Signal that About is ready
  useEffect(() => {
    if (onReady) {
      onReady();
    }
  }, [onReady]);

  return (
    <div className="p-4">
      {/* Navigation Breadcrumb / Back Button */}
      {currentFileName !== HOME_FILE && (
        <button
          onClick={goHome}
          className="mb-4 px-3 py-1 text-sm bg-blue-600 text-white rounded hover:bg-blue-700 transition-colors"
        >
          &larr; Back to Home
        </button>
      )}

      {/* The Viewer */}
      <MarkdownViewer 
        file={fileUrl} 
        onLinkClick={handleLinkClick} 
        theme="light" 
      />
    </div>
  );
};

export default About;