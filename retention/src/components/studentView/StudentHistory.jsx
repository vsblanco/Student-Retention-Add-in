// Timestamp: 2025-10-02 04:37 PM | Version: 3.0.0
import React from 'react';

function StudentHistory({ history }) {
  if (!Array.isArray(history) || history.length === 0) {
    return (
      <div id="history-content">
        <ul className="space-y-4">
          <li className="p-3 bg-gray-100 rounded-lg shadow-sm relative">
            <p className="text-sm text-gray-800">No history found for this student.</p>
          </li>
        </ul>
      </div>
    );
  }

  return (
    <div id="history-content">
      <ul className="space-y-4">
        {history.map((entry, index) => {
          // Always use bg-gray-100 as the default background for comments (matches "No history found" style)
          const bgClass =
            entry.tag === "Contacted"
              ? "bg-yellow-100"
              : "bg-gray-200";
          const tagClass =
            entry.tag === "Contacted"
              ? "px-2 py-0.5 font-semibold rounded-full bg-yellow-200 text-yellow-800"
              : "px-2 py-0.5 font-semibold rounded-full bg-blue-100 text-blue-800";
          return (
            <li
              key={index}
              className={`p-3 rounded-lg shadow-sm relative ${bgClass}`}
              data-row-index={entry.studentId || index}
            >
              <p className="text-sm text-gray-800">{entry.comment}</p>
              <div className="text-xs text-gray-500 mt-2 pt-2 border-t border-gray-200 flex justify-between items-center">
                <div className="flex items-center gap-2">
                  {entry.tag && (
                    <span className={tagClass}>
                      {entry.tag}
                    </span>
                  )}
                  <span className="font-medium">{entry.createdBy}</span>
                </div>
                <span>
                  {entry.timestamp}
                </span>
              </div>
            </li>
          );
        })}
      </ul>
    </div>
  );
}

export default StudentHistory;

