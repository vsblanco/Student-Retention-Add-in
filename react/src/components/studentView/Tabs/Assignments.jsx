// 2025-12-06 13:35 EST - Version 2.5 - Excel whole numbers default to 11:59 PM
import React, { useMemo, useState } from 'react';
import { Clipboard } from 'lucide-react';
import { formatExcelDate } from '../../utility/Conversion';

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
  gradebook: ['Gradebook', 'gradeBookLink', 'Grade Book'],
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

  // compute sorted assignments by due date (ascending)
  // Also filters out assignments with no title or 'N/A'
  const sortedAssignments = useMemo(() => {
    if (!assignments || !assignments.length) return [];

    const monthMap = {
      jan: 0, feb: 1, mar: 2, apr: 3, may: 4, jun: 5,
      jul: 6, aug: 7, sep: 8, sept: 8, oct: 9, nov: 10, dec: 11
    };

    const parseDueToTimestamp = (val) => {
      if (!val) return Number.POSITIVE_INFINITY;

      // 1. Attempt Excel Date conversion if value is numeric
      // "45981" or 45981 would pass !isNaN
      if (!isNaN(val) && String(val).trim() !== '') {
        try {
          let numVal = Number(val);
          // If integer (no time fraction), default to 11:59 PM (0.9993055556)
          if (Number.isInteger(numVal)) {
            numVal += 0.9993055556;
          }

          const converted = formatExcelDate(numVal);
          if (converted) {
            // Assume utility returns a Date object or valid date string
            const dt = new Date(converted);
            if (!isNaN(dt.getTime())) {
              return dt.getTime();
            }
          }
        } catch (e) {
          // ignore error and fall through to string parsing
          console.warn("Excel date parse failed, falling back to string parse", e);
        }
      }

      // 2. Fallback to custom string parsing (Canvas/LMS style)
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

    // FILTER: Exclude assignments with no title or title "N/A"
    const filtered = assignments.filter(a => {
        const titleVal = a[map.title];
        // 1. Check if title exists (not null, undefined, or empty string)
        if (!titleVal) return false;
        // 2. Check if title is explicitly "N/A" (case insensitive)
        if (String(titleVal).trim().toUpperCase() === 'N/A') return false;
        
        return true;
    });

    const copy = [...filtered];
    copy.sort((a, b) => {
      const ta = parseDueToTimestamp(a[map.dueDate]);
      const tb = parseDueToTimestamp(b[map.dueDate]);
      return ta - tb;
    });
    return copy;
  }, [assignments, map]);

  // feedback state for copy button
  const [copied, setCopied] = useState(false);

  // helper: try to query clipboard-write permission; return true if safe to attempt writeText
  const requestClipboardPermission = async () => {
    try {
      if (!navigator.permissions || !navigator.permissions.query) {
        // Permissions API not supported — proceed and let writeText trigger any prompt
        console.log('Permissions API not available — will attempt clipboard write (may prompt).');
        return true;
      }
      // note: some browsers may not support 'clipboard-write' in query; guard it
      const perm = await navigator.permissions.query({ name: 'clipboard-write' });
      // states: 'granted', 'prompt', 'denied'
      if (perm.state === 'denied') {
        console.log('Clipboard permission: denied');
        return false;
      }
      // 'granted' or 'prompt' -> okay to attempt writeText (prompt will appear on write)
      console.log('Clipboard permission:', perm.state);
      return true;
    } catch (e) {
      // any exception => can't determine; attempt writeText (best-effort)
      console.log('Could not determine clipboard permission (error), will attempt write.'); // eslint-disable-line no-console
      return true;
    }
  };

  // copy all assignmentLink values as bullet points
  const copyAssignmentLinks = async () => {
    try {
      const items = (sortedAssignments || []).map(a => {
        const title = a[map.title] || '';
        const link = a[map.assignmentLink] || a[map.link] || '';
        return { title, link };
      }).filter(it => !!it.link);

      if (items.length === 0) {
        // still provide brief feedback if nothing to copy
        setCopied(true);
        setTimeout(() => setCopied(false), 1200);
        return;
      }

      const clean = (s) => String(s || '').replace(/\s+/g, ' ').trim();

      // format bullets as: "- Title: link"
      const bullets = items
        .map(it => `- ${clean(it.title) || 'Untitled'}: ${clean(it.link)}`)
        .join('\n');

      // Ask/check permission before attempting clipboard write.
      const canAttemptClipboard = await requestClipboardPermission();
      let copiedOk = false;
      if (canAttemptClipboard && navigator && navigator.clipboard && typeof navigator.clipboard.writeText === 'function') {
        try {
          await navigator.clipboard.writeText(bullets);
          copiedOk = true;
          console.log('Copied to clipboard via Clipboard API');
        } catch (err) {
          // writeText failed (could be Permissions Policy, cross-origin iframe, etc.)
          console.log('Clipboard writeText failed, falling back to textarea method:', err);
          copiedOk = false;
        }
      }

      if (!copiedOk) {
        // If permission was explicitly denied, give a brief hint for enabling it.
        if (!canAttemptClipboard) {
          try {
            // minimal user hint
            window.alert('Clipboard access is denied. Please allow clipboard access in your browser settings or use the manual copy fallback.');
          } catch (e) {
            // ignore alert failures in embedded contexts
          }
        }

        // textarea fallback (best-effort)
        const ta = document.createElement('textarea');
        ta.value = bullets;
        // off-screen
        ta.style.position = 'fixed';
        ta.style.left = '-9999px';
        document.body.appendChild(ta);
        ta.select();
        try {
          document.execCommand('copy');
          console.log('Copied to clipboard via textarea fallback');
        } catch (e) {
          // ignore; best-effort copy
          console.log('Textarea fallback copy failed', e);
        }
        document.body.removeChild(ta);
      }

      setCopied(true);
      setTimeout(() => setCopied(false), 1200);
    } catch (err) {
      // minimal fallback: no toast here to keep changes small
      setCopied(false);
    }
  };

  // Helper to render the due date. Handles Excel numbers by converting them.
  const renderDueDate = (val) => {
    if (val === undefined || val === null || val === '') return '';
    
    // Check if it is a number (Excel date)
    if (!isNaN(val) && String(val).trim() !== '') {
      try {
        let numVal = Number(val);
        // If integer (no time fraction), default to 11:59 PM (0.9993055556)
        if (Number.isInteger(numVal)) {
          numVal += 0.9993055556;
        }
        return formatExcelDate(numVal);
      } catch (e) {
        // If conversion fails, return original
        return val;
      }
    }
    // Return standard string
    return val;
  };

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

          /* Tooltip styles: positioned to the LEFT of the button, vertically centered.
             Starts slightly right (invisible) and when .visible it fades in and slides left. */
          .tooltip {
            position: absolute;
            top: 50%;
            right: calc(100% + 8px); /* place to the left of the button */
            left: auto;
            transform: translateY(-50%) translateX(12px); /* start slightly to the right */
            background-color: rgba(0, 0, 0, 0.85);
            color: white;
            padding: 6px 10px;
            border-radius: 6px;
            font-size: 12px;
            white-space: nowrap;
            opacity: 0;
            transition: opacity 180ms ease-out, transform 220ms cubic-bezier(.2,.8,.2,1);
            pointer-events: none;
            box-shadow: 0 6px 18px rgba(0,0,0,0.18);
          }
          .tooltip.visible {
            opacity: 1;
            transform: translateY(-50%) translateX(-8px); /* slide left into place */
          }
        `}
      </style>

      { /* New header inserted here */ }
      <div className="sticky-header space-y-4 mb-2 pl-2">
        <div className="flex justify-between items-center">
          <h3 className="text-lg font-bold text-gray-800">Assignments</h3>
          <div className="flex items-center">
            <div className="relative">
              <button
                id="copy-assignments-button"
                className="bg-gray-500 text-white w-8 h-8 rounded-full shadow-lg flex items-center justify-center hover:bg-gray-600"
                aria-label="Copy assignments"
                title={copied ? "Copied" : "Copy assignment links"}
                type="button"
                onClick={copyAssignmentLinks}
              >
                <Clipboard className="h-4 w-4" />
              </button>
              {copied && (
                <div className="tooltip visible">Copied!</div>
              )}
            </div>
          </div>
        </div>
      </div>

      <div
        className="student-assignments max-h-125 overflow-y-auto grid gap-4 custom-scrollbar"
        style={{
          scrollbarWidth: 'thin',
          scrollbarColor: 'rgba(0,0,0,0.15) rgba(0,0,0,0.03)'
        }}
      >
        {sortedAssignments && sortedAssignments.length > 0 ? (
          sortedAssignments.map((a, idx) => {
            // treat missing submission status as false or "Missing" string (i.e. missing)
            const submission = Object.prototype.hasOwnProperty.call(a, map.submission)
              ? a[map.submission]
              : false;
            const isMissing = submission === false ||
                             (typeof submission === 'string' && submission.toLowerCase() === 'missing');
            const borderColor = isMissing ? 'border-red-500' : 'border-blue-300';
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
                       <div className={dueStyle}>
                         {/* UPDATED: Uses renderDueDate helper to format Excel numbers */}
                         Due: {renderDueDate(a[map.dueDate])}
                       </div>
                       {isMissing && (
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
           })
        ) : (
          <div className="text-gray-500 pl-2">No assignments found.</div>
        )}
       </div>
    </>
  );
}

export default StudentAssignments;