/*
 * Timestamp: 2026-03-17 00:00:00
 * Description: Master List report wizard component.
 * Matches the look and feel of CreateLDA/LDAManager.
 * Currently shows campus selection - further steps will add processing logic.
 */

import React, { useState, useEffect } from 'react';
import { ArrowLeft, ChevronDown } from 'lucide-react';

const CAMPUS_MAP = {
  11: 'Deland',
  39: 'Kissimmee',
  10: 'Lakeland',
  49: 'Online Studies',
  9: 'Orlando',
  40: 'Pembroke',
  43: 'South Miami',
  47: 'Tampa'
};

export default function MasterListReport({ onBack }) {
  const [selectedCampus, setSelectedCampus] = useState('all');
  const [includeFutureStart, setIncludeFutureStart] = useState(false);

  return (
    <div className="w-full max-w-2xl mx-auto bg-white rounded-2xl shadow-xl shadow-slate-200/60 border border-white overflow-hidden p-6 transition-all duration-300 min-h-[400px]">

      {/* Header Area */}
      <div className="mb-6 flex items-center justify-between">
        <div>
          <h2 className="text-2xl font-bold tracking-tight text-slate-800">
            Master List
          </h2>
          <p className="text-slate-400 text-sm mt-1">
            Configure your settings below to generate the report
          </p>
        </div>
        <button onClick={onBack} className="text-slate-400 hover:text-slate-600 transition-colors" title="Back to Reports">
          <ArrowLeft className="w-5 h-5" />
        </button>
      </div>

      {/* Settings Form */}
      <section className="flex flex-col gap-6 animate-in fade-in slide-in-from-bottom-4 duration-500">

        {/* Campus Selection */}
        <div className="bg-slate-50/80 rounded-xl border border-slate-100/80 p-4">
          <label className="block text-sm font-medium text-slate-700 mb-2">
            Campus
          </label>
          <div className="relative">
            <select
              value={selectedCampus}
              onChange={(e) => setSelectedCampus(e.target.value)}
              className="w-full appearance-none bg-white border border-slate-200 rounded-lg px-4 py-2.5 text-sm text-slate-700 focus:outline-none focus:ring-2 focus:ring-[#145F82]/30 focus:border-[#145F82] transition-colors cursor-pointer"
            >
              <option value="all">All Campuses</option>
              {Object.entries(CAMPUS_MAP)
                .sort(([, a], [, b]) => a.localeCompare(b))
                .map(([id, name]) => (
                  <option key={id} value={id}>{name}</option>
                ))
              }
            </select>
            <ChevronDown className="absolute right-3 top-1/2 -translate-y-1/2 w-4 h-4 text-slate-400 pointer-events-none" />
          </div>
        </div>

        {/* Include Future Start Toggle */}
        <div className="bg-slate-50/80 rounded-xl border border-slate-100/80 p-4 flex items-center justify-between">
          <label htmlFor="future-start-toggle" className="text-sm font-medium text-slate-700 cursor-pointer">
            Include Future Start
          </label>
          <button
            id="future-start-toggle"
            type="button"
            role="switch"
            aria-checked={includeFutureStart}
            onClick={() => setIncludeFutureStart(!includeFutureStart)}
            className={`relative inline-flex h-6 w-11 shrink-0 rounded-full border-2 border-transparent transition-colors duration-200 ease-in-out cursor-pointer focus:outline-none focus:ring-2 focus:ring-[#145F82]/30 ${
              includeFutureStart ? 'bg-[#145F82]' : 'bg-slate-200'
            }`}
          >
            <span
              className={`pointer-events-none inline-block h-5 w-5 rounded-full bg-white shadow transform transition-transform duration-200 ease-in-out ${
                includeFutureStart ? 'translate-x-5' : 'translate-x-0'
              }`}
            />
          </button>
        </div>

      </section>
    </div>
  );
}
