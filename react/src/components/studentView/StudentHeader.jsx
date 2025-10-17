// Timestamp: 2025-10-03 12:20 PM | Version: 1.2.0
import React from 'react';
import { formatName } from '../utility/Conversion';
import BounceAnimation from '../utility/BounceAnimation';
import DaysOutModal from './DaysOutModal';

// Helper function to get initials from a name string
const getInitials = (name) => {
  if (!name || typeof name !== 'string') return '--';
  const parts = name.split(' ');
  if (parts.length === 1) return parts[0].charAt(0).toUpperCase();
  return (parts[0].charAt(0) + parts[parts.length - 1].charAt(0)).toUpperCase();
};

function StudentHeader({ student }) {
  // Use a fallback student object to prevent errors if the prop is null or undefined
  const safeStudent = student || {};

  // Use formatName only if name is "Last, First"
  const studentNameRaw = safeStudent.StudentName || 'Select a Student';
  const studentName = studentNameRaw.includes(',') ? formatName(studentNameRaw) : studentNameRaw;
  const initials = getInitials(studentName);
  const assignedTo = safeStudent.Assigned || 'Unassigned';
  const daysOut = safeStudent.DaysOut ?? '--'; // Using nullish coalescing to allow 0
  // Determine background color for Days Out
  let daysOutBg = 'bg-gray-200';
  let daysOutText = 'text-gray-800';
  let daysOutLabelText = 'text-gray-500';
  if (typeof daysOut === 'number') {
    if (daysOut === 0) {
      daysOutBg = 'bg-green-300'; // Brighter green for 0 days out
      daysOutText = 'text-black';
      daysOutLabelText = 'text-black';
    } else if (daysOut >= 14) {
      daysOutBg = 'bg-red-400';
      daysOutText = 'text-black';
      daysOutLabelText = 'text-black';
    } else if (daysOut >= 10) {
      daysOutBg = 'bg-red-300';
      daysOutText = 'text-black';
      daysOutLabelText = 'text-black';
    } else if (daysOut >= 5) {
      daysOutBg = 'bg-yellow-200';
      daysOutText = 'text-gray-800';
      daysOutLabelText = 'text-gray-500';
    } else {
      daysOutBg = 'bg-green-200';
      daysOutText = 'text-gray-800';
      daysOutLabelText = 'text-gray-500';
    }
  }
  // Parse grade as a number if possible, handling "%" sign
  let grade = safeStudent.Grade ?? 'N/A';
  let parsedGrade = grade;
  if (typeof grade === 'string' && grade.trim() !== '') {
    const match = grade.match(/^(\d+(?:\.\d+)?)\s*%?$/);
    if (match) {
      parsedGrade = Number(match[1]);
    }
  }
  // Determine background color for Grade
  let gradeBg = 'bg-gray-200';
  let gradeText = 'text-gray-800';
  let gradeTextLabel = 'text-gray-500';
  let gradeDisplay = grade;
  if (typeof parsedGrade === 'number' && !isNaN(parsedGrade)) {
    // Round to nearest integer
    const roundedGrade = Math.round(parsedGrade);
    if (roundedGrade >= 90) {
      gradeBg = 'bg-green-300';
      gradeText = 'text-black';
      gradeTextLabel = 'text-black';
    } else if (roundedGrade >= 70) {
      gradeBg = 'bg-green-200';
      gradeText = 'text-gray-800';
      gradeTextLabel = 'text-gray-500';
    } else if (roundedGrade >= 60) {
      gradeBg = 'bg-yellow-200';
      gradeText = 'text-gray-800';
      gradeTextLabel = 'text-gray-500';
    } else {
      gradeBg = 'bg-red-300';
      gradeText = 'text-black';
      gradeTextLabel = 'text-black';
    }
    gradeDisplay = `${roundedGrade}%`;
  }
  
  // --- Gender-based avatar background ---
  const gender = (safeStudent.Gender || '').toLowerCase();
  const avatarBg =
    gender === 'male'
      ? { background: '#1278FF', color: '#FFFFFF' } // light blue, dark text
      : gender === 'female'
        ? { background: '#ed72b5', color: '#FFFFFF' } // pink, dark pink text
        : { background: '#6b7280', color: '#FFFFFF' };   // default gray, white text

  const gradebookUrl = safeStudent.Gradebook;

  // Helper to check if gradebookUrl is a valid URL
  const isValidGradebookUrl = typeof gradebookUrl === "string" && /^https?:\/\/\S+$/i.test(gradebookUrl);

  // --- Handler to open gradebook ---
  const openGradebook = () => {
    if (!isValidGradebookUrl) return;
    console.log("Gradebook URL clicked:", gradebookUrl);
    if (window.Office && window.Office.context && window.Office.context.ui && window.Office.context.ui.openBrowserWindow) {
      window.Office.context.ui.openBrowserWindow(gradebookUrl);
    } else {
      window.open(gradebookUrl, '_blank');
    }
  };

  const [showNoLink, setShowNoLink] = React.useState(false);
  const [showDaysModal, setShowDaysModal] = React.useState(false);
  const [bounce, setBounce] = React.useState(false);

  return (
    <div className="p-4 bg-white border-b border-gray-200">
      {/* BounceAnimation injects bounce CSS */}
      <BounceAnimation />
      <div className="flex items-center justify-between space-x-4 min-w-0">
        {/* Left side: Avatar and Name */}
        <div className="flex items-center space-x-4 min-w-0">
          <button
            type="button"
            className={`w-12 h-12 rounded-full flex items-center justify-center text-xl font-bold shrink-0 focus:outline-none${bounce ? " bounce" : ""}`}
            style={avatarBg}
            onClick={() => {
              setBounce(true);
              setTimeout(() => setBounce(false), 500);
            }}
            aria-label="Bounce avatar"
          >
            {initials}
          </button>
          <div className="min-w-0">
            <h2
              className="text-lg font-bold text-gray-800 break-words"
              style={{
                display: '-webkit-box',
                WebkitLineClamp: 2,
                WebkitBoxOrient: 'vertical',
                overflow: 'hidden'
              }}
            >
              {studentName}
            </h2>
            <span className="px-2 py-0.5 text-xs font-semibold rounded-full bg-gray-200 text-gray-800 mt-1 inline-block truncate max-w-[120px]">
              {assignedTo}
            </span>
          </div>
        </div>
        {/* Right side: Stats */}
        <div className="flex space-x-2 flex-shrink-0">
          <button
            type="button"
            className={`p-2 text-center rounded-lg ${daysOutBg} ${daysOutText} w-20 focus:outline-none`}
            onClick={() => setShowDaysModal(true)}
            aria-label="Show days out details"
            title="Show days out details"
          >
            <div className="text-xl font-bold">{daysOut}</div>
            <div className={`text-xs font-medium uppercase ${daysOutLabelText}`}>
              {daysOut === 0 ? 'Engaged' : daysOut === 1 ? 'Day Out' : 'Days Out'}
            </div>
          </button>
          <button
            type="button"
            className={`p-2 text-center rounded-lg ${gradeBg} ${gradeText} w-20 transition-colors duration-150 border border-gray-300 hover:border-blue-400`}
            style={{
              outline: 'none',
              position: 'relative',
              cursor: isValidGradebookUrl ? 'pointer' : 'not-allowed'
            }}
            onClick={openGradebook}
            disabled={!isValidGradebookUrl}
            title={isValidGradebookUrl ? "Open Gradebook" : "No Gradebook link"}
            onMouseEnter={() => { if (!isValidGradebookUrl) setShowNoLink(true); }}
            onMouseLeave={() => setShowNoLink(false)}
          >
            <div className="text-xl font-bold">{gradeDisplay}</div>
            <div className={`text-xs font-medium uppercase ${gradeTextLabel}`}>Grade</div>
            {!isValidGradebookUrl && showNoLink && (
              <span
                style={{
                  position: 'absolute',
                  top: 2,
                  right: 2,
                  background: '#f87171',
                  color: '#fff',
                  borderRadius: '6px',
                  fontSize: '10px',
                  padding: '2px 6px',
                  fontWeight: 'bold',
                  zIndex: 2,
                  opacity: showNoLink ? 1 : 0,
                  pointerEvents: 'none',
                  transition: 'opacity 0.15s'
                }}
                aria-label="No gradebook link"
              >
                No Link
              </span>
            )}
          </button>
        </div>
      </div>

      {/* Days Out Modal */}
      <DaysOutModal
        isOpen={showDaysModal}
        onClose={() => setShowDaysModal(false)}
        daysOut={typeof daysOut === 'number' ? daysOut : (Number.isFinite(Number(daysOut)) ? Number(daysOut) : null)}
      />
    </div>
  );
}

export default StudentHeader;
