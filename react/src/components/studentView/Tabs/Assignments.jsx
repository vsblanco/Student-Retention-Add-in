import React, { useMemo } from 'react';
import { FileUp } from 'lucide-react';

const cardStyle =
  // made card slightly smaller (smaller padding and radius, base text-sm)
  'bg-gray-100 shadow rounded-md p-0.5 flex flex-col md:flex-row md:items-center md:justify-between text-sm';
const mainRowStyle = 'flex-1 flex items-center justify-between';
const titleRowStyle = 'font-semibold text-sm text-gray-800 flex items-center'; // reduced from text-base
const iconGroupStyle = 'flex gap-1 ml-2'; // reduced gap
const iconButtonStyle = 'p-1 rounded transition flex items-center justify-center';
const submissionActiveStyle = 'bg-blue-500 hover:bg-blue-600 text-white';
const submissionInactiveStyle = 'bg-gray-300 text-gray-600 cursor-not-allowed';
const dueStyle = 'text-xs text-gray-500'; // reduced
const scoreStyle = 'text-xs text-gray-600'; // reduced

export const COLUMN_ALIASES_ASSIGNMENTS = {
  student: ['Student Name', 'Student'],
  title: ['Assignment Title', 'Title', 'Assignment'],
  dueDate: ['Due Date', 'Deadline'],
  score: ['Score', 'Points'],
  submissionLink: ['Submission Link', 'Submission', 'Submit Link'],
  assignmentLink: ['Assignment Link', 'Assignment URL', 'Assignment Page', 'Link'],
  gradebook: ['Gradebook','gradeBookLink'],
  submission: ['submission', 'Submitted', 'Submission Status', 'Is Submitted']
};

function StudentAssignments({ assignments, reload }) {
  // resolve field map by matching aliases to actual object keys (case/space-insensitive)
  const map = useMemo(() => {
    // no data -> fallback to first alias for each canonical field
    if (!assignments || !assignments.length) {
      const out = {};
      for (const [canonical, aliases] of Object.entries(COLUMN_ALIASES_ASSIGNMENTS)) {
        out[canonical] = (aliases && aliases.length) ? aliases[0] : canonical;
      }
      return out;
    }

    const normalize = (s) => String(s || '').toLowerCase().replace(/[^a-z0-9]/g, '');
    const keys = Object.keys(assignments[0]);

    const out = {};
    for (const [canonical, aliases] of Object.entries(COLUMN_ALIASES_ASSIGNMENTS)) {
      // try to find a key that matches any alias (normalized)
      let found = keys.find((k) => aliases.some((a) => normalize(a) === normalize(k)));
      // fallback: try matching canonical name itself (normalized)
      if (!found) {
        found = keys.find((k) => normalize(k) === normalize(canonical));
      }
      // final fallback: first alias
      out[canonical] = found || ((aliases && aliases.length) ? aliases[0] : canonical);
    }
    return out;
  }, [assignments]);

  // Safe check for empty/undefined assignments
  if (!assignments || !assignments.length) {
    return <div className="text-gray-500">No assignments found.</div>;
  }

  // compute sorted assignments by due date (ascending)
  const sortedAssignments = useMemo(() => {
    const monthMap = {
      jan: 0, feb: 1, mar: 2, apr: 3, may: 4, jun: 5,
      jul: 6, aug: 7, sep: 8, sept: 8, oct: 9, nov: 10, dec: 11
    };

    const parseDueToTimestamp = (val) => {
      if (!val) return Number.POSITIVE_INFINITY;
      const s = String(val).trim();
      // match "Oct 21 by 11:59pm" or "October 21 11:59 pm" or "Oct 21"
      const m = s.match(/^([A-Za-z]+)\s+(\d{1,2})(?:\s*(?:by)?\s*([\d:]+\s*[ap]m)?)?/i);
      if (!m) return Number.POSITIVE_INFINITY;
      const mon = m[1].toLowerCase().slice(0,3);
      const day = parseInt(m[2], 10);
      const timePart = m[3] ? m[3].replace(/\s+/g, '').toLowerCase() : null;

      const monthIdx = monthMap[mon];
      if (monthIdx === undefined) return Number.POSITIVE_INFINITY;

      let hours = 23, minutes = 59; // default to end of day if no time
      if (timePart) {
        const t = timePart.match(/(\d{1,2}):?(\d{2})?(am|pm)?/);
        if (t) {
          const hh = parseInt(t[1], 10);
          const mm = t[2] ? parseInt(t[2], 10) : 0;
          const ampm = t[3];
          hours = hh % 12;
          if (ampm === 'pm') hours += 12;
          if (!ampm && hh === 24) hours = 0; // defensive
          minutes = mm;
        }
      }

      const now = new Date();
      let year = now.getFullYear();
      // build date
      const dt = new Date(year, monthIdx, day, hours, minutes);
      // if parsed date is more than 11 months in the past, assume next year (handles year wrap)
      if (dt.getTime() < now.getTime() - 1000 * 60 * 60 * 24 * 330) {
        dt.setFullYear(year + 1);
      }
      return dt.getTime();
    };

    const copy = [...assignments];
    copy.sort((a, b) => {
      const ta = parseDueToTimestamp(a[map.dueDate]);
      const tb = parseDueToTimestamp(b[map.dueDate]);
      return ta - tb;
    });
    return copy;
  }, [assignments, map]);

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
        {sortedAssignments.map((a, idx) => {
          // treat missing submission status as false (i.e. missing)
          const submission = Object.prototype.hasOwnProperty.call(a, map.submission)
            ? a[map.submission]
            : false;
          const borderColor = submission === false ? 'border-red-500' : 'border-blue-300';
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
                     {submission === false && (
                       <div className="mt-1">
                         {a[map.submissionLink] ? (
                           <a
                             href={a[map.submissionLink]}
                             target="_blank"
                             rel="noopener noreferrer"
                             className="rounded-full bg-white-500 px-2 py-0 flex items-center border border-red-500 cursor-pointer"
                             title="View Submission"
                             style={{ width: '60px', justifyContent: 'center' }} // reduced width
                           >
                             <div className="text-red-500 text-xs font-semibold">missing</div>
                           </a>
                         ) : (
                           <div
                             className="rounded-full bg-white-500 px-2 py-0 flex items-center border border-red-500 opacity-50 cursor-not-allowed"
                             title="No Submission"
                             style={{ width: '70px', justifyContent: 'center' }} // reduced width
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
