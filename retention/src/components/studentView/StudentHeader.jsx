// Timestamp: 2025-10-03 11:15 AM | Version: 1.1.0
import React from 'react';

// Helper function to get initials from a name string
const getInitials = (name) => {
  if (!name || typeof name !== 'string') return '--';
  const parts = name.split(' ');
  if (parts.length === 1) return parts[0].charAt(0).toUpperCase();
  return (parts[0].charAt(0) + parts[parts.length - 1].charAt(0)).toUpperCase();
};

function StudentHeader({ student }) {
  // Use a fallback student object to prevent errors if the prop is null
  const safeStudent = student || {};

  const studentName = safeStudent.StudentName || 'Select a Student';
  const initials = getInitials(safeStudent.StudentName);
  const assignedTo = safeStudent.Assigned || 'Unassigned';
  const daysOut = safeStudent.DaysOut ?? '--'; // Using ?? to allow 0 to be displayed
  const grade = safeStudent.Grade || 'N/A';
  
  return (
    <div>
        <div className="p-4 bg-white border-b border-gray-200">
          <div className="flex items-center justify-between space-x-4">
            {/* Left side: Avatar and Name */}
            <div className="flex items-center space-x-4">
              <div className="w-12 h-12 rounded-full bg-gray-500 text-white flex items-center justify-center text-xl font-bold">
                {initials}
              </div>
              <div>
                <h2 className="text-lg font-bold text-gray-800">{studentName}</h2>
                <span className="px-2 py-0.5 text-xs font-semibold rounded-full bg-gray-200 text-gray-800 mt-1 inline-block">
                  {assignedTo}
                </span>
              </div>
            </div>
            
            {/* Right side: Stats */}
            <div className="flex space-x-2">
              <div className="p-2 text-center rounded-lg bg-gray-200 text-gray-800 w-20">
                <div className="text-xl font-bold">{daysOut}</div>
                <div className="text-xs font-medium uppercase text-gray-500">Days Out</div>
              </div>
              <div className="p-2 text-center rounded-lg bg-gray-200 text-gray-800 w-20 transition-colors duration-150">
                <div className="text-xl font-bold">{grade}</div>
                <div className="text-xs font-medium uppercase text-gray-500">Grade</div>
              </div>
            </div>
          </div>
        </div>
    </div>
  );
}

export default StudentHeader;

