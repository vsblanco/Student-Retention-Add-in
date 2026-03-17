/*
 * Timestamp: 2026-03-17 00:00:00
 * Description: Report Generation side panel component.
 * Opens when the "Generate Report" ribbon button is clicked.
 * Currently displays the header only - further steps will add report generation functionality.
 */

import React, { useEffect } from 'react';
import { FileSpreadsheet } from 'lucide-react';

export default function ReportGeneration({ user, onReady }) {
  useEffect(() => {
    // Signal to App.jsx that this component is ready
    if (onReady) {
      onReady();
    }
  }, [onReady]);

  return (
    <div className="flex flex-col h-screen bg-gray-50">
      {/* Header */}
      <div className="flex items-center gap-3 px-4 py-4 bg-white border-b border-gray-200 shadow-sm">
        <div className="flex items-center justify-center w-9 h-9 rounded-lg bg-blue-100 text-blue-600">
          <FileSpreadsheet size={20} />
        </div>
        <div>
          <h1 className="text-lg font-semibold text-gray-900">Generate Report</h1>
          <p className="text-xs text-gray-500">Create and export retention reports</p>
        </div>
      </div>

      {/* Content area - placeholder for future steps */}
      <div className="flex-1 flex items-center justify-center p-6">
        <p className="text-sm text-gray-400">Report generation options coming soon.</p>
      </div>
    </div>
  );
}
