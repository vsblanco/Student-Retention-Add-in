// Timestamp: 2025-10-03 12:20 PM | Version: 1.2.0
import React from 'react';

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

  const studentName = safeStudent.StudentName || 'Select a Student';
  const initials = getInitials(safeStudent.StudentName);
  const assignedTo = safeStudent.Assigned || 'Unassigned';
  const daysOut = safeStudent.DaysOut ?? '--'; // Using nullish coalescing to allow 0
  // Determine background color for Days Out
  let daysOutBg = 'bg-gray-200';
  let daysOutText = 'text-gray-800';
  let daysOutLabelText = 'text-gray-500';
  if (typeof daysOut === 'number') {
    if (daysOut >= 14) {
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
    if (parsedGrade >= 90) {
      gradeBg = 'bg-green-300';
      gradeText = 'text-black';
      gradeTextLabel = 'text-black';
    } else if (parsedGrade >= 70) {
      gradeBg = 'bg-green-200';
      gradeText = 'text-gray-800';
      gradeTextLabel = 'text-gray-500';
    } else if (parsedGrade >= 60) {
      gradeBg = 'bg-yellow-200';
      gradeText = 'text-gray-800';
      gradeTextLabel = 'text-gray-500';
    } else {
      gradeBg = 'bg-red-300';
      gradeText = 'text-black';
      gradeTextLabel = 'text-black';
    }
    gradeDisplay = `${parsedGrade}%`;
  }
  
  return (
    <div className="p-4 bg-white border-b border-gray-200">
      <div className="flex items-center justify-between space-x-4">
        {/* Left side: Avatar and Name */}
        <div className="flex items-center space-x-4">
          <div className="w-12 h-12 rounded-full bg-gray-500 text-white flex items-center justify-center text-xl font-bold shrink-0 border-2">
            {initials}
          </div>
          <div>
            <h2 className="text-lg font-bold text-gray-800 truncate">{studentName}</h2>
            <span className="px-2 py-0.5 text-xs font-semibold rounded-full bg-gray-200 text-gray-800 mt-1 inline-block">
              {assignedTo}
            </span>
          </div>
        </div>
        
        {/* Right side: Stats */}
        <div className="flex space-x-2">
          <div className={`p-2 text-center rounded-lg ${daysOutBg} ${daysOutText} w-20`}>
            <div className="text-xl font-bold">{daysOut}</div>
            <div className={`text-xs font-medium uppercase ${daysOutLabelText}`}>Days Out</div>
          </div>
          <div className={`p-2 text-center rounded-lg ${gradeBg} ${gradeText} w-20 transition-colors duration-150`}>
            <div className="text-xl font-bold">{gradeDisplay}</div>
            <div className={`text-xs font-medium uppercase ${gradeTextLabel}`}>Grade</div>
          </div>
        </div>
      </div>
    </div>
  );
}

export default StudentHeader;
