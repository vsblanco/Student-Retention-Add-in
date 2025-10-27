import React, { useState, useEffect, useRef } from 'react';
import Comment, { COMMENT_TAGS, CommentSkeleton } from '../Parts/Comment';
import NewComment from '../Modal/NewCommentModal';
import { formatExcelDate, normalizeKeys, formatTimestamp } from '../../utility/Conversion';
import { /* insertRow, editRow, */ } from '../../utility/ExcelAPI'; // removed direct insert/edit usage
import { addComment, deleteComment, generateCommentID } from '../../utility/EditStudentHistory';
import { Folder, FolderOpen } from 'lucide-react';
import { toast } from 'react-toastify';

// Module-level global loading flag + event bus so other modules can trigger History to show skeletons.
let _globalHistoryLoading = false;
const _globalHistoryLoadingBus = new EventTarget();
export function setHistoryLoading(enable = true) {
  _globalHistoryLoading = !!enable;
  _globalHistoryLoadingBus.dispatchEvent(new CustomEvent('change', { detail: _globalHistoryLoading }));
}
export function getHistoryLoading() {
  return _globalHistoryLoading;
}

// Add styles constant
const styles = `
  @keyframes fadeInDrop {
    from {
      opacity: 0;
      transform: translateY(-24px);
    }
    to {
      opacity: 1;
      transform: translateY(0);
    }
  }
  .animate-fadein {
    animation: fadeInDrop 0.4s cubic-bezier(0.4,0,0.2,1);
  }
  /* See-through scrollbar styles */
  #history-content {
    scrollbar-width: thin; /* Firefox */
    scrollbar-color: rgba(0,0,0,0.15) rgba(0,0,0,0.03);
    position: relative; /* allow overlay absolute positioning */
  }
  #history-content::-webkit-scrollbar {
    width: 8px;
    background: transparent;
  }
  #history-content::-webkit-scrollbar-thumb {
    background: rgba(0,0,0,0.15);
    border-radius: 8px;
  }
  #history-content::-webkit-scrollbar-track {
    background: rgba(0,0,0,0.03);
  }

  /* Skeleton overlay that sits above comments and fades out
     Remove padding so layout/spacing exactly matches the normal comment list */
  .skeleton-overlay {
    position: absolute;
    inset: 0;
    z-index: 20;
    background: rgba(255,255,255,0.9);
    display: flex;
    flex-direction: column;
    padding: 0; /* keep items aligned with comments */
    /* Allow skeleton content to overfill the overlay; do NOT show overlay scrollbars */
    overflow: visible;
    transition: opacity 400ms cubic-bezier(0.4,0,0.2,1), visibility 400ms;
  }
`;

function getMonthFromTimestamp(ts) {
  // Try to parse Excel date or ISO string
  if (!ts) return null;
  // If ts is a number, treat as Excel serial
  if (!isNaN(ts)) {
    const excelEpoch = new Date(Date.UTC(1899, 11, 30));
    const date = new Date(excelEpoch.getTime() + ts * 86400000);
    return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}`;
  }
  // If ts is ISO string
  const date = new Date(ts);
  if (!isNaN(date.getTime())) {
    return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}`;
  }
  return null;
}

function getCurrentMonth() {
  const now = new Date();
  return `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}`;
}

function getMonthYearLabel(monthStr) {
  // monthStr: "YYYY-MM"
  const [year, month] = monthStr.split('-');
  const now = new Date();
  const currentYear = now.getFullYear();
  const monthNames = [
    "January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December"
  ];
  const monthIdx = parseInt(month, 10) - 1;
  if (parseInt(year, 10) === currentYear) {
    return monthNames[monthIdx];
  }
  return `${monthNames[monthIdx]} ${year}`;
}

// add helper to parse timestamp to milliseconds
function parseTimestampMs(ts) {
  if (ts == null) return 0;
  // numeric (Excel serial) or numeric string
  if (!isNaN(ts)) {
    const n = Number(ts);
    const excelEpoch = new Date(Date.UTC(1899, 11, 30));
    return excelEpoch.getTime() + n * 86400000;
  }
  const d = new Date(ts);
  return isNaN(d.getTime()) ? 0 : d.getTime();
}

function StudentHistory({ history, student, reload }) {
  console.log('Rendering StudentHistory with history:', history, 'and student:', student);
  // Local state for history to allow adding new comments
  const [localHistory, setLocalHistory] = useState(Array.isArray(history) ? history : []);
  // External/global loading flag: initialized from the exported global state and updated via bus
  const [externalLoading, setExternalLoading] = useState(getHistoryLoading());
  // Background processing ready flag: remains false while normalization/sorting/grouping runs
  const [ready, setReady] = useState(false);
  // Processed data (prepared in background)
  const [processed, setProcessed] = useState({
    pinnedComments: [],
    currentMonthComments: [],
    monthGroups: {},
    sortedMonths: []
  });
  // isLoading true while parent hasn't provided history OR external loading was triggered
  const isLoading = externalLoading || history == null || !ready;

  // When loading starts, immediately clear any previously-processed comments/folders
  // so the old list cannot flash while skeletons are shown.
  useEffect(() => {
    if (isLoading) {
      // NOTE: keep processed comments intact so they render behind the skeleton overlay.
      // Previously this cleared processed state which removed comments while skeletons displayed.
      // We preserve collapsedFolders so the UI behind the overlay remains consistent.
      // ...no-op to intentionally keep comments visible behind skeletons...
    }
  }, [isLoading]);

  // UI state (declare early so hooks order is stable)
  const [showSearch, setShowSearch] = useState(false);
  const [searchTerm, setSearchTerm] = useState('');
  const [showNewComment, setShowNewComment] = useState(false);
  // status message removed in favor of react-toastify toasts

  // Collapsed state for folders — initialize from computed months below (stable hook call)
  const [collapsedFolders, setCollapsedFolders] = useState({});
  const collapsedRef = useRef(collapsedFolders);
  useEffect(() => { collapsedRef.current = collapsedFolders; }, [collapsedFolders]);

  // Sync localHistory with history prop if it changes
  useEffect(() => {
    setLocalHistory(Array.isArray(history) ? history : []);
    // When a real history prop arrives, clear any external/global loading flag so component leaves skeleton state.
    if (history != null) {
      // update local external state and global flag
      setExternalLoading(false);
      setHistoryLoading(false);
    }
    // clear/leave toasts as-is when student changes
}, [history]);

// Subscribe to global loading bus so external callers can toggle loading
useEffect(() => {
  const handler = (e) => {
    setExternalLoading(!!e.detail);
  };
  _globalHistoryLoadingBus.addEventListener('change', handler);
  return () => _globalHistoryLoadingBus.removeEventListener('change', handler);
}, []);
 
   // Background processing: normalize, filter, split pinned/unpinned, sort and group by month.
   // This is done off-render so isLoading remains true until the work finishes and order is stable.
   useEffect(() => {
     // When there's no history, ensure ready is false (we'll keep showing skeletons).
     if (!Array.isArray(localHistory) || localHistory.length === 0) {
       setProcessed({
         pinnedComments: [],
         currentMonthComments: [],
         monthGroups: {},
         sortedMonths: []
       });
       setReady(false);
       return;
     }

     setReady(false);
     let idleId;
     const doWork = () => {
       try {
         const normalizedHistory = localHistory.map(entry => normalizeKeys(entry || {}));
         const filteredHistory = normalizedHistory.filter(
           entry =>
             !searchTerm ||
             (entry.comment && entry.comment.toLowerCase().includes(searchTerm.toLowerCase()))
         );

         const isEntryPinned = (entry) => {
           if (!entry.tag) return false;
           const entryTags = entry.tag.split(',').map(t => t.trim());
           return entryTags.some(tagLabel => {
             const tagObj = COMMENT_TAGS.find(t => t.label === tagLabel);
             if (tagObj) {
               if (tagObj.pinned) return true;
               if (Array.isArray(tagObj.subtags)) {
                 if (tagObj.subtags.some(subtag => subtag.pinned)) return true;
               }
             }
             for (const tag of COMMENT_TAGS) {
               if (Array.isArray(tag.subtags)) {
                 const subtagObj = tag.subtags.find(subtag => subtag.label === tagLabel);
                 if (subtagObj && subtagObj.pinned) return true;
               }
             }
             return false;
           });
         };

         const pinnedComments = filteredHistory
           .filter(isEntryPinned)
           .slice()
           .sort((a, b) => parseTimestampMs(b.timestamp) - parseTimestampMs(a.timestamp));

         const unpinnedComments = filteredHistory
           .filter(entry => !isEntryPinned(entry))
           .slice()
           .sort((a, b) => parseTimestampMs(b.timestamp) - parseTimestampMs(a.timestamp));

         const monthGroups = {};
         const currentMonth = getCurrentMonth();
         const currentMonthComments = [];
         unpinnedComments.forEach(entry => {
           const month = getMonthFromTimestamp(entry.timestamp);
           if (month === currentMonth) {
             currentMonthComments.push(entry);
           } else if (month) {
             if (!monthGroups[month]) monthGroups[month] = [];
             monthGroups[month].push(entry);
           }
         });

         const sortedMonths = Object.keys(monthGroups).sort((a, b) => b.localeCompare(a));

         // Merge collapsed state with new months BEFORE flipping ready so UI doesn't render an intermediate state.
         const prev = collapsedRef.current || {};
         const merged = { ...prev };
         sortedMonths.forEach(month => {
           if (!(month in merged)) merged[month] = true;
         });
         // Remove months that no longer exist
         Object.keys(merged).forEach(m => {
           if (!sortedMonths.includes(m)) delete merged[m];
         });
         // Update collapsedFolders and processed together, then mark ready
         setCollapsedFolders(merged);
         setProcessed({
           pinnedComments,
           currentMonthComments,
           monthGroups,
           sortedMonths
         });
         setReady(true);
       } catch (err) {
         // If processing fails for some reason, mark ready so UI can attempt rendering (and show error via toast if needed)
         setCollapsedFolders({});
         setProcessed({
           pinnedComments: [],
           currentMonthComments: [],
           monthGroups: {},
           sortedMonths: []
         });
         setReady(true);
       }
     };

     if (typeof window !== 'undefined' && 'requestIdleCallback' in window) {
       idleId = window.requestIdleCallback(doWork, { timeout: 500 });
     } else {
       idleId = setTimeout(doWork, 0);
     }

     return () => {
       if (typeof window !== 'undefined' && 'cancelIdleCallback' in window && idleId && typeof idleId === 'number') {
         window.cancelIdleCallback(idleId);
       } else if (idleId) {
         clearTimeout(idleId);
       }
     };
   }, [localHistory, searchTerm]);
 
   function toggleFolder(month) {
     setCollapsedFolders(prev => ({
       ...prev,
       [month]: !prev[month]
     }));
   }

   // Wrapper that delegates to the shared addComment (from StudentView)
   // Keeps UI responsive by updating localHistory immediately; addComment handles the actual sheet insert.
   async function addCommentToHistory(comment, tag = '') {
     if (!comment) return false;
     try {
       // Try to provide student id/name to addComment if available
       const studentId = (student && (student.ID ?? student.Id ?? student.id)) ?? null;
       const studentName = (student && (student.Student ?? student.StudentName ?? student.Name)) ?? null;

       // Persist via shared addComment (actual sheet insert)
       await addComment(String(comment), tag, undefined, studentId, studentName);
       await reload(); // reload history from parent to ensure sync
       toast.success('Comment saved');
       return true;
     } catch (err) {
       // Roll back commentPreview update on error
       setLocalHistory(prev => (Array.isArray(prev) ? prev.filter(e => e !== commentPreviewEntry) : []));
       toast.error('Failed to save comment');
       return false;
     }
   }

   async function deleteCommentFromHistory(commentID) {
     await deleteComment(commentID);
     await reload();
     toast.error('Comment deleted');
   }

   async function saveCommentFromHistory(commentID) {
     await reload();
     toast.success('Comment changes saved');
   }
 
   // add refs to support long-loading toast behavior
   const isLoadingRef = useRef(externalLoading || history == null || !ready);
   const longLoadTimerRef = useRef(null);
   const longLoadToastRef = useRef(null);

   // keep isLoadingRef up-to-date so the timeout callback can read current state
   useEffect(() => {
     isLoadingRef.current = externalLoading || history == null || !ready;
   }, [externalLoading, history, ready]);

   // Show a single toast if loading persists longer than 10 seconds
   useEffect(() => {
     const isLoadingNow = externalLoading || history == null || !ready;
     if (isLoadingNow) {
       // start timer if not already started
       if (!longLoadTimerRef.current) {
         longLoadTimerRef.current = setTimeout(() => {
           // check the ref for current loading state
           if (isLoadingRef.current && !longLoadToastRef.current) {
             longLoadToastRef.current = toast.warn('Uh oh. This is taking longer than usual', { autoClose: 6000 });
           }
         }, 3000); // 3 seconds
       }
     } else {
       // clear any pending timer and reset toast ref (do not dismiss existing toasts automatically)
       if (longLoadTimerRef.current) {
         clearTimeout(longLoadTimerRef.current);
         longLoadTimerRef.current = null;
       }
       longLoadToastRef.current = null;
     }

     // cleanup on unmount
     return () => {
       if (longLoadTimerRef.current) {
         clearTimeout(longLoadTimerRef.current);
         longLoadTimerRef.current = null;
       }
     };
   }, [externalLoading, history, ready]);

   return (
     <div>
       <style>{styles}</style>
       {/* History Header */}
       <div className="sticky-header space-y-4">
         <div className="flex justify-between items-center">
           <h3 className="text-lg font-bold text-gray-800">History</h3>
           <div className="flex items-center space-x-2">
             <button
               id="search-history-button"
               className="bg-gray-600 text-white w-8 h-8 rounded-full shadow-lg flex items-center justify-center hover:bg-gray-700"
               onClick={() => setShowSearch(v => !v)}
               aria-label="Search history"
             >
               <svg xmlns="http://www.w3.org/2000/svg" className="h-5 w-5" fill="none" viewBox="0 0 24 24" stroke="currentColor"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M21 21l-6-6m2-5a7 7 0 11-14 0 7 7 0 0114 0z"></path></svg>
             </button>
             <button
               id="add-comment-button"
               className="bg-blue-600 text-white w-8 h-8 rounded-full shadow-lg flex items-center justify-center hover:bg-blue-700"
               onClick={() => setShowNewComment(v => !v)}
               aria-label="Add comment"
             >
               <svg xmlns="http://www.w3.org/2000/svg" className="h-5 w-5" fill="none" viewBox="0 0 24 24" stroke="currentColor"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M12 6v6m0 0v6m0-6h6m-6 0H6"></path></svg>
             </button>
           </div>
         </div>

         <div
           id="search-container"
           className={`${showSearch ? '' : 'hidden'} space-y-2`}
         >
           <div id="tag-filter-container" className="flex flex-wrap items-center gap-2 pb-2 border-b border-gray-200">
             
           </div>
           <div className="relative">
             <input
               type="text"
               id="search-input"
               className="w-full p-2 pl-8 border rounded-md"
               placeholder="Search comments..."
               value={searchTerm}
               onChange={e => setSearchTerm(e.target.value)}
               autoFocus={showSearch}
             />
             <div className="absolute inset-y-0 left-0 pl-2 flex items-center pointer-events-none">
               <svg className="h-5 w-5 text-gray-400" xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" stroke="currentColor"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M21 21l-6-6m2-5a7 7 0 11-14 0 7 7 0 0114 0z"></path></svg>
             </div>
             <button
               id="clear-search-button"
               className="absolute inset-y-0 right-0 pr-3 flex items-center text-gray-500 hover:text-gray-700"
               type="button"
               onClick={() => setSearchTerm('')}
               aria-label="Clear search"
             >
               ×
             </button>
           </div>
         </div>

         {/* Move new comment box to NewComment component */}
         <NewComment
           show={showNewComment}
           onClose={() => setShowNewComment(false)}
           addCommentToHistory={addCommentToHistory}
         />

         {/* status toasts are handled by the app-level ToastContainer in App.jsx */}
       </div>
       {/* End History Header */}
 
       <div
         id="history-content"
         className="overflow-y-auto"
         style={{
           height: showNewComment || showSearch
             ? 'calc(105vh - 460px)'
             : 'calc(100vh - 260px)'
         }}
       >
         {/* Always render comments (or empty message) — skeletons are shown in a top overlay */}
         <ul className="space-y-4">
           {(!Array.isArray(localHistory) || localHistory.length === 0) ? (
             <li className="p-3 bg-gray-100 rounded-lg shadow-sm relative">
               <p className="text-sm text-gray-800">No history found for this student.</p>
             </li>
           ) : (
             <>
               {/* Pinned comments first (regardless of timestamp) */}
               {processed.pinnedComments.map((entry, index) => (
                 <Comment
                   key={`pinned-${index}`}
                   entry={entry}
                   searchTerm={searchTerm}
                   index={index}
                   delete={deleteCommentFromHistory}
                   save={saveCommentFromHistory}
                 />
               ))}
               {/* Current month comments (not in a folder, only unpinned) */}
               {[...processed.currentMonthComments].map((entry, idx) => (
                 <Comment
                   key={`currentmonth-${idx}`}
                   entry={entry}
                   searchTerm={searchTerm}
                   index={idx}
                   delete={deleteCommentFromHistory}
                   save={saveCommentFromHistory}
                 />
               ))}
               {/* Previous months in collapsible folders (only unpinned) */}
               {(processed.sortedMonths || []).map(month => (
                 <li key={month} className="bg-gray-50 rounded-lg shadow-sm p-2">
                   <div
                     className="flex items-center font-semibold text-gray-700 mb-2 cursor-pointer select-none"
                     onClick={() => toggleFolder(month)}
                     style={{ userSelect: 'none' }}
                     aria-label={collapsedFolders[month] ? "Expand" : "Collapse"}
                     tabIndex={0}
                     role="button"
                     onKeyDown={e => {
                       if (e.key === 'Enter' || e.key === ' ') toggleFolder(month);
                     }}
                   >
                     <span className="mr-2 text-gray-500" style={{ fontSize: '1.2em', lineHeight: '1' }}>
                       {collapsedFolders[month]
                         ? <Folder size={20} strokeWidth={2} />
                         : <FolderOpen size={20} strokeWidth={2} />
                       }
                     </span>
                     {getMonthYearLabel(month)}
                   </div>
                   {!collapsedFolders[month] && (
                     <ul className="space-y-2">
                       {[...(processed.monthGroups[month] || [])].map((entry, idx) => (
                         <Comment
                           key={`month-${month}-${idx}`}
                           entry={entry}
                           searchTerm={searchTerm}
                           index={idx}
                           delete={deleteCommentFromHistory}
                           save={saveCommentFromHistory}
                         />
                       ))}
                     </ul>
                   )}
                 </li>
               ))}
             </>
           )}
         </ul>

         {/* Skeleton overlay sits above the comments and fades away when loading completes */}
         <div
           className="skeleton-overlay"
           style={{
             opacity: isLoading ? 1 : 0,
             pointerEvents: isLoading ? 'auto' : 'none',
             visibility: isLoading ? 'visible' : 'visible' /* keep in DOM for smooth fade; pointerEvents handles interaction */
           }}
           aria-hidden={!isLoading}
         >
           <ul className="space-y-4">
             {Array.from({ length: 5 }).map((_, i) => (
               <CommentSkeleton key={`skeleton-${i}`} />
             ))}
           </ul>
         </div>
       </div>
     </div>
   );
 }
 export default StudentHistory;


