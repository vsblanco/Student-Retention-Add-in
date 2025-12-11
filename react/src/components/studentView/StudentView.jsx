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
  // You can add more aliases for other columns here
};

// add lazy SSO so its module (and any HTML it might insert) is only loaded when needed
const SSO = lazy(() => import('../utility/SSO.jsx'));

// Helper to detect if an Excel address is in the first row (Header)
const isHeaderRowAddress = (address) => {
  if (!address) return false;
  const cleanAddress = address.includes('!') ? address.split('!')[1] : address;
  if (cleanAddress === '1:1') return true;
  return /^[A-Z]+1(:[A-Z]+1)?$/.test(cleanAddress);
};

function StudentView() {
	// Changed default to 'details' to be safe in case History sheet is missing
	const [activeTab, setActiveTab] = useState('details'); 
	const [historyData, setHistory] = useState([]); 
  	const [assignmentData, setAssignments] = useState([]); 
	const [activeStudentState, setActiveStudentState] = useState(activeStudent);
	const [currentUserName, setCurrentUserName] = useState(null);
  
    // State to track if specific sheets exist in the workbook
    const [availableTabs, setAvailableTabs] = useState({
      history: true,    // default to true while checking
      assignments: true // default to true while checking
    });

	// add a ref used elsewhere in the component
	const sessionUserRef = useRef(null);

    // Check for sheet existence on mount
    useEffect(() => {
      const checkSheets = async () => {
        try {
          await Excel.run(async (context) => {
            const sheets = context.workbook.worksheets;
            context.load(sheets, 'items/name');
            await context.sync();
            
            const sheetNames = sheets.items.map(s => s.name);
            
            const hasHistory = sheetNames.includes('Student History');
            // Check for either "Missing Assignments" OR "Assignments"
            const hasAssignments = sheetNames.some(name => 
              ['Missing Assignments', 'Assignments'].includes(name)
            );

            setAvailableTabs({
              history: hasHistory,
              assignments: hasAssignments
            });
            
            console.log('Sheet existence check:', { hasHistory, hasAssignments });
          });
        } catch (error) {
          console.warn('Failed to check sheet existence, defaulting to all visible:', error);
        }
      };

      checkSheets();
    }, []);

    // Effect: If the active tab is set to a tab that doesn't exist, switch to Details
    useEffect(() => {
      if (activeTab === 'history' && !availableTabs.history) {
        setActiveTab('details');
      }
      if (activeTab === 'assignments' && !availableTabs.assignments) {
        setActiveTab('details');
      }
    }, [availableTabs, activeTab]);


	const loadHistory = () => {
    // Guard: Do not attempt load if we don't have a valid ID or if header row
    if (!activeStudentState || !activeStudentState.ID || activeStudentState.ID === 'Student ID') return;
    
    // Guard: Do not attempt to load if History sheet is missing
    if (!availableTabs.history) return;

    console.log('Loading history for', activeStudentState.StudentName);

    setHistoryLoading(true);
    loadSheetCache(activeStudentState.ID)
      .then((cached) => {
        if (cached) {
          setHistory(cached);
        }
      })
      .catch((cacheErr) => {
        console.warn('loadSheetCache failed:', cacheErr);
      })
      .then(() => {
        return loadSheet('Student History', 'StudentNumber', activeStudentState.ID);
      })
      .then((res) => {
        if (res && res.data) {
          setHistory(res.data);
        }
      })
      .catch((err) => {
        console.error('Failed to load Student History sheet:', err);
      });
	};

  const loadAssignments = () => {
    if (!activeStudentState || !activeStudentState.Gradebook) return;
    // Guard: Do not attempt load if Assignment sheet is missing
    if (!availableTabs.assignments) return;

    console.log('Loading assignments for', activeStudentState.StudentName);
		loadSheet('Missing Assignments', 'gradeBookLink', activeStudentState.Gradebook)
			.then((res) => {
				setAssignments(res.data);
			})
			.catch((err) => {
				// Fallback: try loading "Assignments" if "Missing Assignments" failed (optional)
				console.error('Failed to load Student Assignments sheet:', err);
			});
	};

	const renderActiveTab = () => {
    // If blank row or header selected
    if (!activeStudentState || !activeStudentState.ID) {
      return (
        <div className="p-8 text-center text-gray-500">
          <p>Please select a student row to view details.</p>
        </div>
      );
    }

		switch (activeTab) {
		case 'history':
			return availableTabs.history ? 
        <StudentHistory history={historyData} student={activeStudentState} reload={loadHistory} /> : 
        <StudentDetails student={activeStudentState} />;
		case 'assignments':
			return availableTabs.assignments ? 
        <StudentAssignments assignments={assignmentData} reload={loadAssignments} /> : 
        <StudentDetails student={activeStudentState} />;
		case 'details':
		default:
			return <StudentDetails student={activeStudentState} />
		}
	};

  // Register Excel selection-changed handler
  useEffect(() => {
	let handlerRef = null;
	(async () => {
	  try {
		handlerRef = await onSelectionChanged(({ address, values, data }) => {
		  console.log('Excel selection changed:', { address, data });

      // 1. Check if the selection is the Header Row (Row 1)
      if (isHeaderRowAddress(address)) {
        console.log('Header row selected - clearing active student view.');
        setActiveStudentState({}); 
        return; 
      }

		  setActiveStudentState(prev => ({ ...prev, ...data }));
		}, COLUMN_ALIASES);

		// On initial load
		try {
		  const sel = await getSelectedRange(COLUMN_ALIASES);
		  if (sel && sel.success) {
		    const initialRow = sel.singleRow || (Array.isArray(sel.rows) && sel.rows[0]) || null;
		    if (initialRow) {
          // Data check: If the values look like headers, ignore them.
          if (
            initialRow.ID === 'Student ID' || 
            initialRow.StudentName === 'Student Name' || 
            initialRow.Student === 'Student'
          ) {
             console.log('Initial selection is header row - ignoring.');
             setActiveStudentState({});
          } else {
            setActiveStudentState(prev => ({ ...prev, ...initialRow }));
            try {
              loadCache(initialRow);
            } catch (e) {
              console.warn('loadCache initial invocation failed', e);
            }
          }
		    }
		  }
		} catch (gErr) {
		  console.warn('getSelectedRange failed to initialize selection:', gErr);
		}
	  } catch (err) {
		console.error('Failed to register Excel selection handler:', err);
	  }
	})();

	return () => {
	  if (handlerRef && typeof handlerRef.remove === 'function') {
		try {
		  handlerRef.remove();
		} catch (err) {
		  console.warn('Failed to remove selection handler:', err);
		}
	  }
	};
  }, []);

  // Register Excel cell-change handler
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

  useEffect(() => {
    try {
      const cached = window.localStorage.getItem('ssoUserName') || window.localStorage.getItem('SSO_USER');
      if (cached) {
        setCurrentUserName(cached);
        sessionUserRef.current = cached;
        return;
      }
      if (window.SSO && typeof window.SSO.getUserName === 'function') {
        const n = window.SSO.getUserName();
        if (n) {
          setCurrentUserName(n);
          sessionUserRef.current = n;
        }
      }
    } catch (_) { /* ignore */ }
  }, []);

  useEffect(() => {
	if (!activeStudentState?.ID) return;
	// Only load if the tabs are available
    if (availableTabs.history) loadHistory();
	if (availableTabs.assignments) loadAssignments();
	// eslint-disable-next-line react-hooks/exhaustive-deps
  }, [activeStudentState?.ID]);

  if (!currentUserName) {
    return (
      <div className="studentview-outer">
        <Suspense fallback={null}>
          <SSO onNameSelect={setCurrentUserName} />
        </Suspense>
      </div>
    );
  }

  return (
		<div className="studentview-outer">
			<StudentHeader student={activeStudentState} />
			<div className="studentview-tabs">
				<button
					type="button"
					className={`studentview-tab ${activeTab === 'details' ? 'active' : ''}`}
					onClick={() => setActiveTab('details')}
				>
					Details
				</button>
				
                {/* Conditionally Render History Tab Button */}
                {availableTabs.history && (
                    <button
                        type="button"
                        className={`studentview-tab ${activeTab === 'history' ? 'active' : ''}`}
                        onClick={() => {
                            loadHistory();
                            setActiveTab('history');
                        }}
                    >
                        History
                    </button>
                )}

                {/* Conditionally Render Assignments Tab Button */}
                {availableTabs.assignments && (
                    <button
                        type="button"
                        className={`studentview-tab ${activeTab === 'assignments' ? 'active' : ''}`}
                        onClick={() => {
                            loadAssignments();
                            setActiveTab('assignments');
                        }}
                    >
                        Assignments
                    </button>
                )}
			</div>

			<div className="studentview-tab-content">
				{renderActiveTab()}
			</div>
		</div>
	);
}

StudentView.displayName = 'StudentView';

export default React.memo(StudentView);