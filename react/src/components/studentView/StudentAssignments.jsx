import React from 'react';
import { FileUp, BookOpen } from 'lucide-react';

// Expects props.student.Assignments to be an array of assignment objects
// Each assignment: { title, dueDate, score, submissionLink, assignmentLink }

const cardStyle =
  'bg-white shadow rounded-lg p-4 flex flex-col md:flex-row md:items-center md:justify-between';
const mainRowStyle = 'flex-1 flex items-center justify-between';
const titleRowStyle = 'font-semibold text-base text-gray-800 flex items-center';
const iconGroupStyle = 'flex gap-2 ml-2';
const iconButtonStyle = 'p-1 rounded transition flex items-center justify-center';
const submissionActiveStyle = 'bg-blue-500 hover:bg-blue-600 text-white';
const submissionInactiveStyle = 'bg-gray-300 text-gray-600 cursor-not-allowed';
const assignmentActiveStyle = 'bg-green-500 hover:bg-green-600 text-white';
const assignmentInactiveStyle = 'bg-gray-300 text-gray-600 cursor-not-allowed';
const dueStyle = 'text-sm text-gray-500';
const scoreStyle = 'text-sm text-gray-600';

function StudentAssignments({ student }) {
  const assignments = student?.Assignments || [];

  if (!assignments.length) {
    return <div className="text-gray-500">No assignments found.</div>;
  }

  return (
    <div className="student-assignments grid gap-4">
      {assignments.map((a, idx) => (
        <div key={idx} className={cardStyle}>
          <div className={mainRowStyle}>
            <div>
              <div className={titleRowStyle}>
                {/* You can add line-clamp here if needed */}
                {a.title}
                <div className={iconGroupStyle}>
                  {a.submissionLink ? (
                    <a
                      href={a.submissionLink}
                      target="_blank"
                      rel="noopener noreferrer"
                      className={`${iconButtonStyle} ${submissionActiveStyle}`}
                      title="View Submission"
                    >
                      <FileUp className="w-4 h-4" />
                    </a>
                  ) : (
                    <button
                      className={`${iconButtonStyle} ${submissionInactiveStyle}`}
                      disabled
                      title="No Submission"
                    >
                      <FileUp className="w-4 h-4 opacity-50" />
                    </button>
                  )}
                  {a.assignmentLink ? (
                    <a
                      href={a.assignmentLink}
                      target="_blank"
                      rel="noopener noreferrer"
                      className={`${iconButtonStyle} ${assignmentActiveStyle}`}
                      title="View Assignment"
                    >
                      <BookOpen className="w-4 h-4" />
                    </a>
                  ) : (
                    <button
                      className={`${iconButtonStyle} ${assignmentInactiveStyle}`}
                      disabled
                      title="No Assignment"
                    >
                      <BookOpen className="w-4 h-4 opacity-50" />
                    </button>
                  )}
                </div>
              </div>
              <div className={dueStyle}>Due: {a.dueDate}</div>
              <div className={scoreStyle}>Score: {a.score}</div>
            </div>
          </div>
        </div>
      ))}
    </div>
  );
}

export default StudentAssignments;
