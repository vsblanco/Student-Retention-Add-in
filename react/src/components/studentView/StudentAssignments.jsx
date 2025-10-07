import React from 'react';
import { FileUp, BookOpen } from 'lucide-react';

const cardStyle =
  'bg-gray-100 shadow rounded-lg p-4 flex flex-col md:flex-row md:items-center md:justify-between';
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

function StudentAssignments({ student, fieldMap }) {
  const assignments = student?.Assignments || [];
  // Default field mapping
  const map = fieldMap || {
    title: 'title',
    dueDate: 'dueDate',
    score: 'score',
    submissionLink: 'submissionLink',
    assignmentLink: 'assignmentLink',
    submission: 'submission' // new field
  };

  if (!assignments.length) {
    return <div className="text-gray-500">No assignments found.</div>;
  }

  return (
    <div className="student-assignments max-h-96 overflow-y-auto grid gap-4">
      {assignments.map((a, idx) => {
        const borderColor =
          a[map.submission] === false
            ? "border-red-500"
            : "border-blue-300";
        return (
          <div key={idx} className={`${cardStyle} group border-l-4 ${borderColor} pl-4`}>
            <div className={`${mainRowStyle} w-full`}>
              <div className="flex-1">
                <div className={titleRowStyle}>
                  {a[map.title]}
                </div>
                <div className={dueStyle}>Due: {a[map.dueDate]}</div>
                <div className={scoreStyle + " flex items-center"}>
                  Score: {a[map.score] === '' || a[map.score] === undefined || a[map.score] === null ? 0 : a[map.score]}
                  {a[map.submission] === false && (
                    <span
                      className="ml-2 inline-block"
                      dir="ltr"
                    >
                      <div className="rounded-full bg-white-500 px-3 py-0.2 flex items-center border border-red-500">
                        <div className="text-red-500 text-xs font-semibold">missing</div>
                      </div>
                    </span>
                  )}
                </div>
              </div>
              <div className={`${iconGroupStyle} flex-shrink-0`}>
                {a[map.submissionLink] ? (
                  <a
                    href={a[map.submissionLink]}
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
                {a[map.assignmentLink] ? (
                  <a
                    href={a[map.assignmentLink]}
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
          </div>
        );
      })}
    </div>
  );
}

export default StudentAssignments;
