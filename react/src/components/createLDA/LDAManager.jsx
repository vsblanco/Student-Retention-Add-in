/*
 * Timestamp: 2026-01-26 00:00:00
 * Description: Fully functional LDA Manager. Integrates with ldaProcessor.js to run the real Excel generation logic.
 * Update: Added batch progress display for large datasets (6000+ students). Shows progress bar under
 *         "Reading Master List" and "Formatting LDA Table" steps when multiple batches are being processed.
 */

import React, { useState, useEffect } from 'react';
import { Info, CheckCircle2, Circle, Loader2, ArrowLeft, AlertCircle, ChevronRight } from 'lucide-react';
import { createLDA, detectCampuses } from './ldaProcessor';

// --- CONFIGURATION: Steps matching the processor logic ---
const PROCESS_STEPS = [
  { id: 'validate', label: 'Validating Workbook Settings' },
  { id: 'read', label: 'Reading Master List' },
  { id: 'filter', label: 'Filtering by Days Out' },
  { id: 'failing', label: 'Filtering by Grades' },
  { id: 'createSheet', label: 'Creating Sheet' },
  { id: 'tags', label: 'Applying LDA & DNC Tags' },
  { id: 'format', label: 'Formatting LDA Table' },
  { id: 'finalize', label: 'Finalizing Report' },
];

export default function CreateLDAManager({ onReady } = {}) {
  // State for the settings
  const [ldaSettings, setLdaSettings] = useState({
    daysOut: 5,
    includeFailingList: false,
    includeLDATag: true,
    includeDNCTag: true,
    sheetNameMode: 'date',
  });

  // State for View Management: 'settings' | 'processing' | 'done' | 'error'
  const [view, setView] = useState('settings');
  const [settingsView, setSettingsView] = useState('main');
  const [errorMessage, setErrorMessage] = useState('');

  // State for Progress Tracking: { [stepId]: 'pending' | 'active' | 'completed' }
  const [stepStatus, setStepStatus] = useState({});

  // State for Batch Progress (for large datasets)
  // { current, total, phase, tableName } where phase is 'writing' | 'formatting'
  const [batchProgress, setBatchProgress] = useState(null);

  // State for Campus Detection (when campus mode is active)
  const [detectedCampuses, setDetectedCampuses] = useState([]);
  const [campusLoading, setCampusLoading] = useState(false);

  // State for Multi-Campus Processing Progress
  // { campusName: 'pending' | 'active' | 'completed' }
  const [campusStatuses, setCampusStatuses] = useState({});
  const [isMultiCampus, setIsMultiCampus] = useState(false);

  // Load workbook settings on mount
  useEffect(() => {
    let isMounted = true;

    const loadSettings = () => {
      if (typeof window !== 'undefined' && window.Office && Office.context && Office.context.document && Office.context.document.settings) {
        Office.context.document.settings.refreshAsync((result) => {
          if (result.status === Office.AsyncResultStatus.Succeeded && isMounted) {
            const wb = Office.context.document.settings.get('workbookSettings');
            if (wb && typeof wb === 'object') {
              setLdaSettings(prev => ({
                daysOut: (wb.daysOut !== undefined && wb.daysOut !== null) ? Number(wb.daysOut) : prev.daysOut,
                includeFailingList: (wb.includeFailingList !== undefined) ? !!wb.includeFailingList : prev.includeFailingList,
                includeLDATag: (wb.includeLDATag !== undefined) ? !!wb.includeLDATag : ((wb.includeLdatTag !== undefined) ? !!wb.includeLdatTag : prev.includeLDATag),
                includeDNCTag: (wb.includeDNCTag !== undefined) ? !!wb.includeDNCTag : ((wb.includeDncTag !== undefined) ? !!wb.includeDncTag : prev.includeDNCTag),
                sheetNameMode: wb.sheetNameMode || prev.sheetNameMode,
              }));
            }
          }
        });
      }
    };

    loadSettings();

    if (typeof window !== 'undefined' && window.Office && Office.context && Office.context.document && Office.context.document.settings) {
      Office.context.document.settings.addHandlerAsync(Office.EventType.SettingsChanged, loadSettings);
    }

    return () => {
      isMounted = false;
      if (typeof window !== 'undefined' && window.Office && Office.context && Office.context.document && Office.context.document.settings) {
        Office.context.document.settings.removeHandlerAsync(Office.EventType.SettingsChanged, loadSettings);
      }
    };
  }, []);

  // Signal that LDAManager is ready
  useEffect(() => {
    if (onReady) {
      onReady();
    }
  }, [onReady]);

  // Detect campuses when campus mode is toggled on
  useEffect(() => {
    if (ldaSettings.sheetNameMode === 'campus') {
      setCampusLoading(true);
      detectCampuses()
        .then(campuses => {
          setDetectedCampuses(campuses);
          setCampusLoading(false);
        })
        .catch(() => {
          setDetectedCampuses([]);
          setCampusLoading(false);
        });
    } else {
      setDetectedCampuses([]);
    }
  }, [ldaSettings.sheetNameMode]);

  const handleSettingChange = (key, value) => {
    setLdaSettings((prev) => ({ ...prev, [key]: value }));
  };

  // --- REAL LOGIC: Trigger the Processor ---
  const handleCreateLDA = async () => {
    console.log('Starting LDA Creation Process with:', ldaSettings);

    const multiCampus = ldaSettings.sheetNameMode === 'campus' && detectedCampuses.length > 1;

    // 1. Switch View
    setView('processing');
    setErrorMessage('');
    setBatchProgress(null);
    setIsMultiCampus(multiCampus);

    // 2. Reset Steps
    const initialStatus = {};
    PROCESS_STEPS.forEach(s => initialStatus[s.id] = 'pending');
    setStepStatus(initialStatus);

    // Initialize campus statuses if multi-campus
    if (multiCampus) {
      const initial = {};
      detectedCampuses.forEach(c => initial[c] = 'pending');
      setCampusStatuses(initial);
    } else {
      setCampusStatuses({});
    }

    try {
        // 3. Call the imported processor function
        await createLDA(
            ldaSettings,
            // Step progress callback
            (stepId, status) => {
                setStepStatus(prev => ({
                    ...prev,
                    [stepId]: status
                }));
                // Clear batch progress when moving to a new step
                if (status === 'active') {
                    setBatchProgress(null);
                }
            },
            // Batch progress callback (for large datasets)
            (current, total, phase, tableName) => {
                setBatchProgress({ current, total, phase, tableName });
            },
            // Campus progress callback (for multi-campus mode)
            multiCampus ? (campusName, campusIndex, totalCampuses, status) => {
                setCampusStatuses(prev => ({
                    ...prev,
                    [campusName]: status
                }));
            } : null
        );

        // 4. Success
        setBatchProgress(null);
        setView('done');

    } catch (error) {
        console.error("LDA Failed:", error);
        setErrorMessage(error.message || "An unexpected error occurred.");
        setBatchProgress(null);
        setView('error');
    }
  };

  const handleReset = () => {
    setView('settings');
    setStepStatus({});
    setErrorMessage('');
    setBatchProgress(null);
    setCampusStatuses({});
    setIsMultiCampus(false);
  };

  return (
    <div className="w-full max-w-2xl mx-auto bg-white rounded-2xl shadow-xl shadow-slate-200/60 border border-white overflow-hidden p-6 transition-all duration-300 min-h-[400px]">
      
      {/* Header Area */}
      <div className="mb-6 flex items-center justify-between">
        <div>
            <h2 className={`text-2xl font-bold tracking-tight ${view === 'error' ? 'text-red-600' : 'text-slate-800'}`}>
              {view === 'settings' ? 'Create LDA' : 
               view === 'done' ? 'Complete' : 
               view === 'error' ? 'Error' :
               'Processing...'}
            </h2>
            <p className="text-slate-400 text-sm mt-1">
                {view === 'settings' 
                  ? 'Configure your settings below to generate the report' 
                  : (view === 'done' ? 'Your LDA report has been generated.' : 
                     view === 'error' ? 'Something went wrong.' :
                     'Please wait while we generate your report.')}
            </p>
        </div>
        {/* Back button (only show if not in settings view) */}
        {view !== 'settings' && (
             <button onClick={handleReset} className="text-slate-400 hover:text-slate-600 transition-colors" title="Back to Settings">
                <ArrowLeft className="w-5 h-5" />
             </button>
        )}
      </div>

      <div className="relative">
        {/* VIEW 1: Settings Form */}
        {view === 'settings' && (
          <section className="flex flex-col gap-6 animate-in fade-in slide-in-from-bottom-4 duration-500">
            <LDASettings
              settings={ldaSettings}
              onSettingChange={handleSettingChange}
              settingsView={settingsView}
              setSettingsView={setSettingsView}
            />

            {settingsView === 'main' && (
              <div className="pt-4 border-t border-slate-100 mt-2">
                <button
                  type="button"
                  className="w-full sm:w-auto bg-[#145F82] hover:bg-[#0f4b66] text-white font-medium px-6 py-2.5 rounded-full shadow-lg shadow-[#145F82]/20 transition-all duration-200 hover:-translate-y-0.5 flex items-center justify-center gap-2"
                  onClick={handleCreateLDA}
                >
                  Create LDA
                </button>

                {/* Campus Detection List - fades in when campus mode is active */}
                <div
                  className={`transition-all duration-500 ease-in-out overflow-hidden ${
                    ldaSettings.sheetNameMode === 'campus'
                      ? 'opacity-100 max-h-96 mt-4'
                      : 'opacity-0 max-h-0 mt-0'
                  }`}
                >
                  {campusLoading ? (
                    <div className="flex items-center gap-2 text-slate-400 text-sm p-3">
                      <Loader2 className="w-4 h-4 animate-spin" />
                      <span>Scanning for campuses...</span>
                    </div>
                  ) : detectedCampuses.length > 1 ? (
                    <div className="bg-slate-50/80 rounded-xl border border-slate-100/80 p-4">
                      <div className="flex items-center gap-2 mb-3">
                        <div className="w-5 h-5 rounded-md bg-[#145F82]/10 flex items-center justify-center">
                          <span className="text-[#145F82] text-xs font-bold">{detectedCampuses.length}</span>
                        </div>
                        <span className="text-sm font-medium text-slate-700">
                          Campuses Detected
                        </span>
                      </div>
                      <div className="space-y-1.5">
                        {detectedCampuses.map(campus => (
                          <div key={campus} className="flex items-center gap-2.5 text-sm text-slate-600 pl-1">
                            <div className="w-1.5 h-1.5 rounded-full bg-[#145F82]/50" />
                            <span>{campus}</span>
                          </div>
                        ))}
                      </div>
                      <p className="text-xs text-slate-400 mt-3 pt-3 border-t border-slate-100">
                        {detectedCampuses.length} separate LDA sheets will be created
                      </p>
                    </div>
                  ) : detectedCampuses.length === 1 ? (
                    <div className="text-sm text-slate-500 p-3 bg-slate-50/80 rounded-xl border border-slate-100/80">
                      1 campus found: <span className="font-medium text-slate-700">{detectedCampuses[0]}</span>
                    </div>
                  ) : ldaSettings.sheetNameMode === 'campus' && !campusLoading ? (
                    <div className="text-sm text-slate-400 italic p-3">
                      No Campus column found in Master List
                    </div>
                  ) : null}
                </div>
              </div>
            )}
          </section>
        )}

        {/* VIEW 2: Processing Checklist */}
        {(view === 'processing' || view === 'done' || view === 'error') && (
           <section className="flex flex-col gap-4 animate-in fade-in slide-in-from-right-8 duration-500">
              <div className="bg-slate-50 rounded-xl border border-slate-100 p-4 space-y-3">
                  {/* Render steps - for multi-campus, show global steps then campus list */}
                  {(isMultiCampus
                    ? PROCESS_STEPS.filter(s => ['validate', 'read', 'filter', 'failing', 'tags'].includes(s.id))
                    : PROCESS_STEPS
                  ).map((step) => {
                      const status = stepStatus[step.id] || 'pending';
                      const showBatchProgress = (step.id === 'format' || step.id === 'read') &&
                          status === 'active' && batchProgress && batchProgress.total > 1;

                      return (
                          <div key={step.id}>
                              <div className="flex items-center gap-3">
                                  <div className="w-6 flex justify-center shrink-0">
                                      {status === 'pending' && <Circle className="w-4 h-4 text-slate-300" />}
                                      {status === 'active' && <Loader2 className="w-5 h-5 text-[#145F82] animate-spin" />}
                                      {status === 'completed' && <CheckCircle2 className="w-5 h-5 text-emerald-500 animate-in zoom-in duration-300" />}
                                  </div>
                                  <div className={`flex-1 text-sm font-medium transition-colors duration-300 ${
                                      status === 'pending' ? 'text-slate-400' :
                                      status === 'active' ? 'text-slate-800' : 'text-emerald-700'
                                  }`}>
                                      {step.label}
                                  </div>
                              </div>

                              {showBatchProgress && (
                                  <div className="ml-9 mt-2 animate-in fade-in slide-in-from-top-2 duration-300">
                                      <div className="flex items-center justify-between text-xs text-slate-500 mb-1">
                                          <span>
                                              {batchProgress.phase === 'reading' ? 'Reading' :
                                               batchProgress.phase === 'writing' ? 'Writing' : 'Formatting'}{' '}
                                              {batchProgress.tableName === 'Master_List' ? 'Master List' :
                                               batchProgress.tableName === 'LDA_Table' ? 'LDA' : 'Failing'} data
                                          </span>
                                          <span className="font-medium">
                                              {batchProgress.current} / {batchProgress.total}
                                          </span>
                                      </div>
                                      <div className="h-1.5 bg-slate-200 rounded-full overflow-hidden">
                                          <div
                                              className="h-full bg-[#145F82] rounded-full transition-all duration-300 ease-out"
                                              style={{ width: `${(batchProgress.current / batchProgress.total) * 100}%` }}
                                          />
                                      </div>
                                  </div>
                              )}
                          </div>
                      );
                  })}

                  {/* Multi-Campus Progress Section */}
                  {isMultiCampus && Object.keys(campusStatuses).length > 0 && (
                      <div className="mt-1 pt-3 border-t border-slate-200/60">
                          <div className="flex items-center gap-2 mb-2.5">
                              <span className="text-xs font-semibold text-slate-500 uppercase tracking-wider">
                                  Campus Reports
                              </span>
                              <span className="text-xs text-slate-400">
                                  ({Object.values(campusStatuses).filter(s => s === 'completed').length} / {Object.keys(campusStatuses).length})
                              </span>
                          </div>
                          <div className="space-y-2">
                              {Object.entries(campusStatuses).map(([campusName, status]) => (
                                  <div key={campusName} className="flex items-center gap-3 animate-in fade-in duration-300">
                                      <div className="w-6 flex justify-center shrink-0">
                                          {status === 'pending' && <Circle className="w-3.5 h-3.5 text-slate-300" />}
                                          {status === 'active' && <Loader2 className="w-4 h-4 text-[#145F82] animate-spin" />}
                                          {status === 'completed' && <CheckCircle2 className="w-4 h-4 text-emerald-500 animate-in zoom-in duration-300" />}
                                      </div>
                                      <div className={`flex-1 text-sm transition-colors duration-300 ${
                                          status === 'pending' ? 'text-slate-400' :
                                          status === 'active' ? 'text-slate-700 font-medium' : 'text-emerald-600'
                                      }`}>
                                          {campusName}
                                      </div>
                                  </div>
                              ))}
                          </div>
                      </div>
                  )}
              </div>

              {view === 'error' && (
                  <div className="pt-2 animate-in fade-in slide-in-from-bottom-2 duration-500">
                      <div className="bg-red-50 text-red-700 px-4 py-3 rounded-lg text-sm font-medium flex items-center gap-3 border border-red-100">
                          <AlertCircle className="w-5 h-5 shrink-0" />
                          <span>{errorMessage}</span>
                      </div>
                      <button
                        onClick={handleReset}
                        className="mt-4 w-full text-slate-500 hover:text-slate-700 text-sm font-medium hover:underline"
                      >
                        Try Again
                      </button>
                  </div>
              )}

              {view === 'done' && (
                  <div className="pt-2 animate-in fade-in slide-in-from-bottom-2 duration-500">
                      <div className="bg-emerald-50 text-emerald-700 px-4 py-3 rounded-lg text-sm font-medium flex items-center justify-center border border-emerald-100">
                          {isMultiCampus
                            ? `${Object.keys(campusStatuses).length} LDA Sheets Generated Successfully!`
                            : 'LDA Generated Successfully!'
                          }
                      </div>
                      <button
                        onClick={handleReset}
                        className="mt-4 w-full text-slate-500 hover:text-slate-700 text-sm font-medium hover:underline"
                      >
                        Start Over
                      </button>
                  </div>
              )}
           </section>
        )}
      </div>
    </div>
  );
}

// --- Sub-Components ---

function LDASettings({ settings, onSettingChange, settingsView, setSettingsView }) {

  const handleToggle = (key) => {
    onSettingChange(key, !settings[key]);
  };

  const handleInputChange = (key, value) => {
    onSettingChange(key, value);
  };

  if (settingsView === 'tags') {
    return (
      <div className="flex flex-col gap-4 w-full animate-in fade-in slide-in-from-right-4 duration-300">
        <button
          type="button"
          onClick={() => setSettingsView('main')}
          className="flex items-center gap-1 text-slate-400 hover:text-slate-600 text-sm font-medium transition-colors w-fit"
        >
          <ArrowLeft className="w-4 h-4" />
          Back
        </button>

        <ToggleRow
          label="Include LDA Tag"
          isOn={settings.includeLDATag}
          onToggle={() => handleToggle('includeLDATag')}
        />

        <ToggleRow
          label="Include DNC Tag"
          isOn={settings.includeDNCTag}
          onToggle={() => handleToggle('includeDNCTag')}
        />
      </div>
    );
  }

  return (
    <div className="flex flex-col gap-4 w-full">
      <div className="flex items-center justify-between p-3 bg-slate-50/50 rounded-xl border border-slate-100/50 hover:border-slate-200 transition-colors">
        <div className="flex items-center gap-2">
          <span className="text-slate-700 font-medium text-sm">Sheet Name</span>
          <Info className="w-4 h-4 text-slate-400 cursor-help hover:text-slate-600" />
        </div>
        <div className="relative flex bg-slate-200 rounded-full p-0.5" style={{ width: 130 }}>
          <button
            type="button"
            onClick={() => handleInputChange('sheetNameMode', 'date')}
            className={`relative z-10 flex-1 py-1 text-xs font-medium rounded-full text-center transition-colors duration-200 ${
              settings.sheetNameMode === 'date'
                ? 'text-white'
                : 'text-slate-500 hover:text-slate-700'
            }`}
          >
            Date
          </button>
          <button
            type="button"
            onClick={() => handleInputChange('sheetNameMode', 'campus')}
            className={`relative z-10 flex-1 py-1 text-xs font-medium rounded-full text-center transition-colors duration-200 ${
              settings.sheetNameMode === 'campus'
                ? 'text-white'
                : 'text-slate-500 hover:text-slate-700'
            }`}
          >
            Campus
          </button>
          <span
            className="absolute top-0.5 bottom-0.5 rounded-full bg-[#145F82] transition-all duration-200"
            style={{
              width: 'calc(50% - 2px)',
              left: settings.sheetNameMode === 'date' ? '2px' : 'calc(50%)',
            }}
          />
        </div>
      </div>

      <div className="flex items-center justify-between p-3 bg-slate-50/50 rounded-xl border border-slate-100/50 hover:border-slate-200 transition-colors">
        <div className="flex items-center gap-2">
          <label htmlFor="daysOut" className="text-slate-700 font-medium text-sm">
            Days Out
          </label>
          <div className="group relative">
             <Info className="w-4 h-4 text-slate-400 cursor-help hover:text-slate-600" />
          </div>
        </div>
        <input
          id="daysOut"
          type="number"
          value={settings.daysOut}
          onChange={(e) => handleInputChange('daysOut', Number(e.target.value))}
          className="w-24 border border-slate-200 bg-white rounded-lg px-3 py-1.5 text-right text-sm focus:outline-none focus:ring-2 focus:ring-[#145F82]/20 focus:border-[#145F82] transition-all"
        />
      </div>

      <ToggleRow
        label="Include Failing List"
        isOn={settings.includeFailingList}
        onToggle={() => handleToggle('includeFailingList')}
      />

      <button
        type="button"
        onClick={() => setSettingsView('tags')}
        className="flex items-center justify-between p-3 bg-slate-50/50 rounded-xl border border-slate-100/50 hover:border-slate-200 transition-colors"
      >
        <div className="flex items-center gap-2">
          <span className="text-slate-700 font-medium text-sm">Tags</span>
          <Info className="w-4 h-4 text-slate-400 cursor-help hover:text-slate-600" />
        </div>
        <ChevronRight className="w-4 h-4 text-slate-400" />
      </button>
    </div>
  );
}

function ToggleRow({ label, isOn, onToggle }) {
  return (
    <div className="flex items-center justify-between p-3 bg-slate-50/50 rounded-xl border border-slate-100/50 hover:border-slate-200 transition-colors">
      <div className="flex items-center gap-2">
        <span className="text-slate-700 font-medium text-sm">{label}</span>
        <Info className="w-4 h-4 text-slate-400 cursor-help hover:text-slate-600" />
      </div>
      
      <button
        type="button"
        onClick={onToggle}
        className={`relative inline-flex h-6 w-11 items-center rounded-full transition-colors duration-200 focus:outline-none focus:ring-2 focus:ring-[#145F82]/20 focus:ring-offset-1 ${
          isOn ? 'bg-[#145F82]' : 'bg-slate-200'
        }`}
      >
        <span
          className={`${
            isOn ? 'translate-x-6' : 'translate-x-1'
          } inline-block h-4 w-4 transform rounded-full bg-white shadow-sm transition-transform duration-200`}
        />
      </button>
    </div>
  );
}