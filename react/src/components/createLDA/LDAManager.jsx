/*
 * Timestamp: 2026-01-26 00:00:00
 * Description: Fully functional LDA Manager. Integrates with ldaProcessor.js to run the real Excel generation logic.
 * Update: Added batch progress display for large datasets (6000+ students). Shows progress bar under
 *         "Reading Master List" and "Formatting LDA Table" steps when multiple batches are being processed.
 */

import React, { useState, useEffect } from 'react';
import { Info, CheckCircle2, Circle, Loader2, ArrowLeft, AlertCircle } from 'lucide-react';
import { createLDA } from './ldaProcessor'; // Import the new logic

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
  });

  // State for View Management: 'settings' | 'processing' | 'done' | 'error'
  const [view, setView] = useState('settings');
  const [errorMessage, setErrorMessage] = useState('');

  // State for Progress Tracking: { [stepId]: 'pending' | 'active' | 'completed' }
  const [stepStatus, setStepStatus] = useState({});

  // State for Batch Progress (for large datasets)
  // { current, total, phase, tableName } where phase is 'writing' | 'formatting'
  const [batchProgress, setBatchProgress] = useState(null);

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

  const handleSettingChange = (key, value) => {
    setLdaSettings((prev) => ({ ...prev, [key]: value }));
  };

  // --- REAL LOGIC: Trigger the Processor ---
  const handleCreateLDA = async () => {
    console.log('Starting LDA Creation Process with:', ldaSettings);

    // 1. Switch View
    setView('processing');
    setErrorMessage('');
    setBatchProgress(null);

    // 2. Reset Steps
    const initialStatus = {};
    PROCESS_STEPS.forEach(s => initialStatus[s.id] = 'pending');
    setStepStatus(initialStatus);

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
            }
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
            />

            <div className="pt-4 border-t border-slate-100 mt-2">
              <button
                type="button"
                className="w-full sm:w-auto bg-[#145F82] hover:bg-[#0f4b66] text-white font-medium px-6 py-2.5 rounded-full shadow-lg shadow-[#145F82]/20 transition-all duration-200 hover:-translate-y-0.5 flex items-center justify-center gap-2"
                onClick={handleCreateLDA}
              >
                Create LDA
              </button>
            </div>
          </section>
        )}

        {/* VIEW 2: Processing Checklist */}
        {(view === 'processing' || view === 'done' || view === 'error') && (
           <section className="flex flex-col gap-4 animate-in fade-in slide-in-from-right-8 duration-500">
              <div className="bg-slate-50 rounded-xl border border-slate-100 p-4 space-y-3">
                  {PROCESS_STEPS.map((step, index) => {
                      const status = stepStatus[step.id] || 'pending';
                      // Show batch progress for 'read' or 'format' steps when processing large datasets
                      const showBatchProgress = (step.id === 'format' || step.id === 'read') &&
                          status === 'active' && batchProgress && batchProgress.total > 1;

                      return (
                          <div key={step.id}>
                              <div className="flex items-center gap-3">
                                  {/* Icon Column */}
                                  <div className="w-6 flex justify-center shrink-0">
                                      {status === 'pending' && <Circle className="w-4 h-4 text-slate-300" />}
                                      {status === 'active' && <Loader2 className="w-5 h-5 text-[#145F82] animate-spin" />}
                                      {status === 'completed' && <CheckCircle2 className="w-5 h-5 text-emerald-500 animate-in zoom-in duration-300" />}
                                  </div>

                                  {/* Text Column */}
                                  <div className={`flex-1 text-sm font-medium transition-colors duration-300 ${
                                      status === 'pending' ? 'text-slate-400' :
                                      status === 'active' ? 'text-slate-800' : 'text-emerald-700'
                                  }`}>
                                      {step.label}
                                  </div>
                              </div>

                              {/* Batch Progress (shown under format step for large datasets) */}
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
                          LDA Generated Successfully!
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

function LDASettings({ settings, onSettingChange }) {
  const handleToggle = (key) => {
    onSettingChange(key, !settings[key]);
  };

  const handleInputChange = (key, value) => {
    onSettingChange(key, value);
  };

  return (
    <div className="flex flex-col gap-4 w-full">
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

function ToggleRow({ label, isOn, onToggle }) {
  return (
    <div className="flex items-center justify-between p-3 bg-slate-50/50 rounded-xl border border-slate-100/50 hover:border-slate-200 transition-colors">
      <div className="flex items-center gap-2">
        <span className="text-slate-700 font-medium text-sm">{label}</span>
        <Info className="w-4 h-4 text-slate-400 cursor-help hover:text-slate-600" />
      </div>
      
      <div className="flex items-center gap-3">
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
        <span className="text-xs font-medium text-slate-500 w-6">
          {isOn ? 'On' : 'Off'}
        </span>
      </div>
    </div>
  );
}