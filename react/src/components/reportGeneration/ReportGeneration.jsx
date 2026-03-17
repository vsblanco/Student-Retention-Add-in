/*
 * Timestamp: 2026-03-17 00:00:00
 * Description: Report Generation side panel component.
 * Shows a list of available reports. Clicking a report opens its wizard.
 * Currently only Master List is active - other reports are greyed out.
 */

import React, { useState, useEffect } from 'react';
import { ChevronRight } from 'lucide-react';
import MasterListReport from './MasterListReport.jsx';

const REPORTS = [
  { id: 'master-list', label: 'Master List', enabled: true },
  { id: 'lda', label: 'LDA', enabled: false },
  { id: 'failing', label: 'Failing', enabled: false },
  { id: 'attendance', label: 'On-Ground Attendance', enabled: false },
  { id: 'deans-presidents', label: "Dean's & President's", enabled: false },
];

export default function ReportGeneration({ user, onReady }) {
  const [activeReport, setActiveReport] = useState(null);

  useEffect(() => {
    if (onReady) {
      onReady();
    }
  }, [onReady]);

  // If a report wizard is open, render it
  if (activeReport === 'master-list') {
    return <MasterListReport onBack={() => setActiveReport(null)} />;
  }

  return (
    <div className="w-full max-w-2xl mx-auto bg-white rounded-2xl shadow-xl shadow-slate-200/60 border border-white overflow-hidden p-6 transition-all duration-300 min-h-[400px]">

      {/* Header Area */}
      <div className="mb-6">
        <h2 className="text-2xl font-bold tracking-tight text-slate-800">
          Generate Report
        </h2>
        <p className="text-slate-400 text-sm mt-1">
          Select a report to get started
        </p>
      </div>

      {/* Report List */}
      <section className="flex flex-col gap-2 animate-in fade-in slide-in-from-bottom-4 duration-500">
        {REPORTS.map((report) => (
            <button
              key={report.id}
              type="button"
              disabled={!report.enabled}
              onClick={() => report.enabled && setActiveReport(report.id)}
              className={`w-full flex items-center gap-3 px-4 py-3 rounded-xl border text-left transition-all duration-200
                ${report.enabled
                  ? 'border-slate-100 bg-white hover:bg-slate-50 hover:border-slate-200 cursor-pointer active:scale-[0.99]'
                  : 'border-slate-50 bg-slate-50/50 cursor-not-allowed'
                }`}
            >
              <span className={`flex-1 text-sm font-medium ${
                report.enabled ? 'text-slate-700' : 'text-slate-300'
              }`}>
                {report.label}
              </span>
              {report.enabled && (
                <ChevronRight className="w-4 h-4 text-slate-300" />
              )}
            </button>
        ))}
      </section>
    </div>
  );
}
