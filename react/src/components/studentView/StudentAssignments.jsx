import React from 'react';
import { FileUp } from 'lucide-react';

const cardStyle =
  'bg-gray-100 shadow rounded-lg p-1 flex flex-col md:flex-row md:items-center md:justify-between';
const mainRowStyle = 'flex-1 flex items-center justify-between';
const titleRowStyle = 'font-semibold text-base text-gray-800 flex items-center';
const iconGroupStyle = 'flex gap-2 ml-2';
const iconButtonStyle = 'p-1 rounded transition flex items-center justify-center';
const submissionActiveStyle = 'bg-blue-500 hover:bg-blue-600 text-white';
const submissionInactiveStyle = 'bg-gray-300 text-gray-600 cursor-not-allowed';
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
    <>
      <style>
        {`
          .student-assignments.custom-scrollbar::-webkit-scrollbar {
            width: 8px;
            background: transparent;
          }
          .student-assignments.custom-scrollbar::-webkit-scrollbar-thumb {
            background: rgba(0,0,0,0.15);
            border-radius: 8px;
          }
          .student-assignments.custom-scrollbar::-webkit-scrollbar-track {
            background: rgba(0,0,0,0.03);
          }
        `}
      </style>
      <div
        className="student-assignments max-h-125 overflow-y-auto grid gap-4 custom-scrollbar"
        style={{
          scrollbarWidth: 'thin',
          scrollbarColor: 'rgba(0,0,0,0.15) rgba(0,0,0,0.03)'
        }}
      >
        {assignments.map((a, idx) => {
          const borderColor =
            a[map.submission] === false
              ? "border-red-500"
              : "border-blue-300";
          return (
            <div key={idx} className={`${cardStyle} group border-l-4 ${borderColor} pl-4`}>
              <div className={`${mainRowStyle} w-full`}>
                <div className="flex-1 flex flex-col justify-between h-full">
                  <div>
                    <div className={titleRowStyle}>
                      {a[map.assignmentLink] ? (
                        <a
                          href={a[map.assignmentLink]}
                          target="_blank"
                          rel="noopener noreferrer"
                          title="View Assignment"
                          style={{ textDecoration: 'none' }}
                        >
                          {a[map.title]}
                        </a>
                      ) : (
                        <span title="No Assignment">
                          {a[map.title]}
                        </span>
                      )}
                    </div>
                    <div className={dueStyle}>Due: {a[map.dueDate]}</div>
                    {a[map.submission] === false && (
                      <div className="mt-1">
                        {a[map.submissionLink] ? (
                          <a
                            href={a[map.submissionLink]}
                            target="_blank"
                            rel="noopener noreferrer"
                            className="rounded-full bg-white-500 px-3 py-0.5 flex items-center border border-red-500 cursor-pointer"
                            title="View Submission"
                            style={{ width: '60px', justifyContent: 'center' }} // fixed width
                          >
                            <div className="text-red-500 text-xs font-semibold">missing</div>
                          </a>
                        ) : (
                          <div
                            className="rounded-full bg-white-500 px-3 py-0.5 flex items-center border border-red-500 opacity-50 cursor-not-allowed"
                            title="No Submission"
                            style={{ width: '90px', justifyContent: 'center' }} // fixed width
                          >
                            <div className="text-red-500 text-xs font-semibold">missing</div>
                          </div>
                        )}
                      </div>
                    )}
                  </div>
                  <div className="flex justify-end mt-2">
                    <div className={scoreStyle + " flex items-center"}>
                      Score: {a[map.score] === '' || a[map.score] === undefined || a[map.score] === null ? 0 : a[map.score]}
                    </div>
                  </div>
                </div>
                <div className={`${iconGroupStyle} flex-shrink-0`}>
                  {/* Removed old view submission button */}
                </div>
              </div>
            </div>
          );
        })}
      </div>
    </>
  );
}


export default StudentAssignments;
