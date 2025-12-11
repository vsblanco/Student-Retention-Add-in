// 2025-12-11 13:15 EST - Version 4.4.0 - Fix premature ready signal race condition
import React, { useState, useEffect, useRef, lazy, Suspense } from 'react';
import './Styling/StudentView.css';
import StudentHeader from './Parts/Header.jsx';
import StudentDetails from './Tabs/Details.jsx';
import StudentHistory, { setHistoryLoading } from './Tabs/History.jsx';
import StudentAssignments from './Tabs/Assignments.jsx';
import { onSelectionChanged, highlightRow, loadSheet, getSelectedRange, onChanged } from '../utility/ExcelAPI.jsx';
import { loadCache, loadSheetCache } from '../utility/Cache.jsx';
import { isOutreachTrigger } from './Tag';
import { addComment } from '../utility/EditStudentHistory.jsx';

/* global Excel */

const activeStudent = {};
export const COLUMN_ALIASES = {
  StudentName: ['Student Name', 'Student'],
  ID: ['Student ID', 'SyStudentID','Student identifier'],
  StudentNumber: ['Student Number'],
  Shift: ['Shift'],
  Instructor: ['Instructor', 'Teacher'],
  ProgVersDescrip: ['Program', 'Program Description', 'ProgVersDescrip'],
  Gender: ['Gender'],
  Phone: ['Phone Number', 'Contact'],
  CreatedBy: ['Created By', 'Author'],
  OtherPhone: ['Other Phone', 'Alt Phone'],
  StudentEmail: ['Email', 'Student Email'],
  PersonalEmail: ['Other Email'],
  Assigned: ['Advisor'],
  AdmissionsRep: ['Admissions Rep', 'Admissions','AdmRep','AmRep'],
  ExpectedStartDate: ['Expected Start Date', 'Start Date','ExpStartDate'],
  Grade: ['Current Grade', 'Grade %', 'Grade', 'Course Grade'],
  LDA: ['Last Date of Attendance', 'LDA', 'course lda', 'CurrentLDA'],
  DaysOut: ['Days Out'],
  Gradebook: ['Gradebook','gradeBookLink'],
  MissingAssignments: ['Missing Assignments', 'Missing'],
  Outreach: ['Outreach', 'Comments', 'Notes', 'Comment'],
  ProfilePicture: ['Profile Picture', 'Photo', 'Avatar']
};

const SSO = lazy(() => import('../utility/SSO.jsx'));

const isHeaderRowAddress = (address) => {
  if (!address) return false;
  const cleanAddress = address.includes('!') ? address.split('!')[1] : address;
  if (cleanAddress === '1:1') return true;
  return /^[A-Z]+1(:[A-Z]+1)?$/.test(cleanAddress);
};

function StudentView({ onReady }) {
    const [activeTab, setActiveTab] = useState('details'); 
    const [historyData, setHistory] = useState([]); 
    const [assignmentData, setAssignments] = useState([]); 
    const [activeStudentState, setActiveStudentState] = useState(activeStudent);
    
    // NEW: Initialize checking state to true so we don't fall through to SSO immediately
    const [isCheckingUser, setIsCheckingUser] = useState(true);
    const [currentUserName, setCurrentUserName] = useState(null);
    
    const [availableTabs, setAvailableTabs] = useState({
      history: true,
      assignments: true
    });

    const sessionUserRef = useRef(null);

    // 1. Check for Sheet Existence
    useEffect(() => {
      const checkSheets = async () => {
        try {
          await Excel.run(async (context) => {
            const sheets = context.workbook.worksheets;
            context.load(sheets, 'items/name');
            await context.sync();
            const sheetNames = sheets.items.map(s => s.name);
            setAvailableTabs({
              history: sheetNames.includes('Student History'),
              assignments: sheetNames.some(name => ['Missing Assignments', 'Assignments'].includes(name))
            });
          });
        } catch (error) {
          console.warn('Failed to check sheet existence', error);
        }
      };
      checkSheets();
    }, []);

    // 2. Tab Auto-Switch Logic
    useEffect(() => {
      if (activeTab === 'history' && !availableTabs.history) setActiveTab('details');
      if (activeTab === 'assignments' && !availableTabs.assignments) setActiveTab('details');
    }, [availableTabs, activeTab]);

    const loadHistory = () => {
        if (!activeStudentState || !activeStudentState.StudentNumber) return;
        if (!availableTabs.history) return;

        console.log('Loading history for', activeStudentState.StudentName, 'using Number:', activeStudentState.StudentNumber);
        setHistoryLoading(true);
        
        loadSheetCache(activeStudentState.StudentNumber)
            .then((cached) => { if (cached) setHistory(cached); })
            .catch(() => {});
            
        loadSheet('Student History', 'StudentNumber', activeStudentState.StudentNumber)
            .then((res) => { if (res && res.data) setHistory(res.data); })
            .catch((err) => console.error('Failed to load Student History sheet:', err));
    };

    const loadAssignments = () => {
        if (!activeStudentState || !activeStudentState.Gradebook) return;
        if (!availableTabs.assignments) return;
        loadSheet('Missing Assignments', 'gradeBookLink', activeStudentState.Gradebook)
            .then((res) => { setAssignments(res.data); })
            .catch((err) => console.error(err));
    };

    // 3. AUTO-LOADER
    useEffect(() => {
        if (!activeStudentState?.StudentNumber) return;
        if (availableTabs.history) loadHistory();
        if (availableTabs.assignments) loadAssignments();
        // eslint-disable-next-line react-hooks/exhaustive-deps
    }, [activeStudentState?.StudentNumber]);

    const renderActiveTab = () => {
        if (!activeStudentState || !activeStudentState.ID) {
            return (
                <div className="p-8 text-center text-gray-500">
                    <p>Please select a student row to view details.</p>
                </div>
            );
        }
        switch (activeTab) {
        case 'history':
            return availableTabs.history ? <StudentHistory history={historyData} student={activeStudentState} reload={loadHistory} /> : <StudentDetails student={activeStudentState} />;
        case 'assignments':
            return availableTabs.assignments ? <StudentAssignments assignments={assignmentData} reload={loadAssignments} /> : <StudentDetails student={activeStudentState} />;
        case 'details':
        default:
            return <StudentDetails student={activeStudentState} />
        }
    };

  // 4. Selection Handler & Initial Load
  useEffect(() => {
    let handlerRef = null;
    (async () => {
      try {
        handlerRef = await onSelectionChanged(({ address, values, data }) => {
          if (isHeaderRowAddress(address)) {
            setActiveStudentState({}); 
            return; 
          }
          setActiveStudentState(prev => ({ ...prev, ...data }));
        }, COLUMN_ALIASES);

        try {
          const sel = await getSelectedRange(COLUMN_ALIASES);
          if (sel && sel.success) {
            const initialRow = sel.singleRow || (Array.isArray(sel.rows) && sel.rows[0]) || null;
            if (initialRow) {
                if (initialRow.ID === 'Student ID' || initialRow.StudentName === 'Student Name') {
                    setActiveStudentState({});
                } else {
                    setActiveStudentState(prev => ({ ...prev, ...initialRow }));
                    try { 
                        // Await cache so we don't lift loader too early
                        await loadCache(initialRow); 
                    } catch (e) { 
                        console.warn('loadCache failed', e); 
                    }
                }
            }
          }
        } catch (gErr) {
          console.warn('getSelectedRange failed to initialize selection:', gErr);
        } finally {
            // UPDATED LOGIC:
            // Only signal ready here if we HAVE a user. 
            // If we are not logged in, we let the SSO block below handle the ready signal.
            if (sessionUserRef.current && onReady) {
                console.log("StudentView Data Load Complete (User logged in) -> Signaling Ready");
                onReady();
            }
        }

      } catch (err) {
        console.error('Failed to register Excel selection handler:', err);
        // Only force ready if we are logged in, otherwise let SSO handle it
        if (sessionUserRef.current && onReady) onReady();
      }
    })();

    return () => {
      if (handlerRef && typeof handlerRef.remove === 'function') {
        handlerRef.remove();
      }
    };
  }, []); 

  // 5. Outreach Handler
  useEffect(() => {
    let changeHandlerRef = null;
    (async () => {
      try {
        changeHandlerRef = await onChanged(
          (changeEvent) => {
            const changes = (changeEvent && Array.isArray(changeEvent.changes)) ? changeEvent.changes : [];
            const matches = changes.map(ch => {
              const rawVal = ch && ch.value;
              const text = (rawVal === undefined || rawVal === null) ? '' : String(rawVal);
              try {
                const out = isOutreachTrigger(text); 
                return { change: ch, text, match: Boolean(out && out.matched), tag: out && out.tag };
              } catch (e) {
                return { change: ch, text, match: false, tag: null };
              }
            });
            matches.forEach(({ change, text, match, tag }) => {
              try {
                if (match) {
                  highlightRow(change.rowIndex, change.colIndex, 9);
                  const tagString = tag ? `${tag}, Outreach` : 'Outreach';
                  addComment(String(text), tagString, undefined, change.otherValues?.ID, change.otherValues?.StudentName);
                } else {
                  addComment(String(text), 'Outreach', undefined, change.otherValues?.ID, change.otherValues?.StudentName);
                }
              } catch (e) {
                console.warn('highlightRow/addComment failed', e);
              }
            });
            const anyMatch = matches.some(m => m.match);
            return anyMatch;
          },
          null,               
          'Outreach',               
          COLUMN_ALIASES,     
          ['StudentName','ID'] 
        );
      } catch (err) {
        console.error('Failed to register Excel cell-change handler:', err);
      }
    })();

    return () => {
      if (changeHandlerRef && typeof changeHandlerRef.remove === 'function') {
        try {
          changeHandlerRef.remove();
        } catch (err) {
          console.warn('Failed to remove cell-change handler:', err);
        }
      }
    };
  }, []);

  // 6. SSO Check (Updated)
  useEffect(() => {
    const checkUser = () => {
        try {
            const cached = window.localStorage.getItem('ssoUserName') || window.localStorage.getItem('SSO_USER');
            if (cached) {
                setCurrentUserName(cached);
                sessionUserRef.current = cached;
            } else if (window.SSO && typeof window.SSO.getUserName === 'function') {
                const n = window.SSO.getUserName();
                if (n) {
                    setCurrentUserName(n);
                    sessionUserRef.current = n;
                }
            }
        } catch (_) { 
            // ignore
        } finally {
            // CRITICAL: We mark checking as done regardless of outcome
            setIsCheckingUser(false);
        }
    };
    checkUser();
  }, []);

  // CONDITION 1: Still checking user?
  // Return null so we don't render SSO or calling onReady prematurely.
  if (isCheckingUser) {
    return null;
  }

  // CONDITION 2: User Check Done, but No User Found?
  // Render SSO and signal ready (so loading screen lifts to reveal login form).
  if (!currentUserName) {
    if (onReady) {
        setTimeout(onReady, 100); 
    }
    return (
      <div className="studentview-outer">
        <Suspense fallback={null}>
          <SSO onNameSelect={setCurrentUserName} />
        </Suspense>
      </div>
    );
  }

  // CONDITION 3: User Logged In
  // Render App. Ready signal is handled by the Excel Selection effect above.
  return (
        <div className="studentview-outer">
            <StudentHeader student={activeStudentState} />
            <div className="studentview-tabs">
                <button type="button" className={`studentview-tab ${activeTab === 'details' ? 'active' : ''}`} onClick={() => setActiveTab('details')}>Details</button>
                {availableTabs.history && ( <button type="button" className={`studentview-tab ${activeTab === 'history' ? 'active' : ''}`} onClick={() => { loadHistory(); setActiveTab('history'); }}>History</button> )}
                {availableTabs.assignments && ( <button type="button" className={`studentview-tab ${activeTab === 'assignments' ? 'active' : ''}`} onClick={() => { loadAssignments(); setActiveTab('assignments'); }}>Assignments</button> )}
            </div>
            <div className="studentview-tab-content">
                {renderActiveTab()}
            </div>
        </div>
    );
}

StudentView.displayName = 'StudentView';
export default React.memo(StudentView);