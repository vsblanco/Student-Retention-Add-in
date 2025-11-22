/* * Timestamp: 2025-11-22 15:00:00 EST
 * Version: 5.0.1
 * Author: Gemini (for Victor)
 * Description: Standalone FileCard component. Features a 'loading' state with a 3-dot pulsing animation.
 * FIX: Replaced external asset import with inline SVG data URI to prevent build errors.
 */

import React from 'react';
import { File, CheckCircle } from 'lucide-react';

// Embedded SVG for CSV icon to avoid external file dependency errors
const csvIcon = `data:image/svg+xml;utf8,<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 32 32" fill="none" stroke="%23f97316" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect x="4" y="4" width="24" height="24" rx="4" fill="%23fff7ed"/><path d="M10 10h12"/><path d="M10 16h12"/><path d="M10 22h8"/></svg>`;

export default function FileCard({ file, rows, type, action, icon, status = 'normal' } = {}) {
    const name = (file && (file.name || file.filename)) || 'Unknown.csv';
    const sizeKB = file && file.size ? Math.round(file.size / 1024) : null;

    // 1. Define styles based on status
    const stateStyles = {
        normal: 'bg-white border-slate-100 shadow-sm hover:shadow-md hover:border-slate-200',
        completed: 'bg-emerald-50/60 border-emerald-200 shadow-sm',
        pending: 'bg-slate-50 border-slate-100 opacity-60 grayscale pointer-events-none',
        // Loading: Looks like pending (bg-slate-50) to fade content, but relative/cursor-wait for the overlay
        loading: 'bg-slate-50 border-slate-100 relative cursor-wait', 
    };

    // 2. Define icon container styles
    const iconStyles = {
        normal: 'bg-slate-50 border-slate-100',
        completed: 'bg-white border-emerald-100',
        pending: 'bg-slate-100 border-slate-200',
        loading: 'bg-slate-100 border-slate-200', // Match pending style
    };

    return (
        <div
            title={name}
            className={`
                w-full group relative
                p-3 rounded-xl border
                transition-all duration-300 ease-in-out
                ${stateStyles[status] || stateStyles.normal}
            `}
        >
            {/* Inject custom keyframes for the dot animation within the component */}
            <style>
                {`
                    @keyframes dot-scale {
                        0%, 100% { transform: scale(0.75); opacity: 0.5; }
                        50% { transform: scale(1.5); opacity: 1; }
                    }
                `}
            </style>

            {/* --- LOADING OVERLAY (3 Dots) --- */}
            {status === 'loading' && (
                <div className="absolute inset-0 z-20 flex items-center justify-center">
                    <div className="bg-white px-3 py-2 rounded-full shadow-sm ring-1 ring-purple-100 animate-in fade-in zoom-in duration-200 flex items-center gap-1">
                        {/* Dot 1 */}
                        <div 
                            className="w-2 h-2 bg-purple-600 rounded-full"
                            style={{ animation: 'dot-scale 1.2s infinite ease-in-out both', animationDelay: '0s' }}
                        />
                        {/* Dot 2 */}
                        <div 
                            className="w-2 h-2 bg-purple-600 rounded-full"
                            style={{ animation: 'dot-scale 1.2s infinite ease-in-out both', animationDelay: '0.2s' }}
                        />
                        {/* Dot 3 */}
                        <div 
                            className="w-2 h-2 bg-purple-600 rounded-full"
                            style={{ animation: 'dot-scale 1.2s infinite ease-in-out both', animationDelay: '0.4s' }}
                        />
                    </div>
                </div>
            )}

            {/* --- INNER CONTENT WRAPPER --- */}
            {/* We fade this out if loading to mimic 'pending', but keep the spinner bright on top */}
            <div className={`flex items-center gap-3 ${status === 'loading' ? 'opacity-50 grayscale blur-[0.5px]' : ''}`}>
                
                {/* --- Icon Column --- */}
                <div className={`
                    w-11 h-11 flex-none flex items-center justify-center 
                    rounded-lg border overflow-hidden relative
                    ${iconStyles[status] || iconStyles.normal}
                `}>
                    
                    {/* Icons: Now shown even in Loading state (to look like pending) */}
                    {icon ? (
                        <img src={icon} alt="import type" className="w-8 h-8 object-contain rounded-md" />
                    ) : /\.csv$/i.test(name) ? (
                        <img src={csvIcon} alt="CSV" className="w-8 h-8 object-contain rounded-md" />
                    ) : (
                        <File size={20} className="text-slate-400 rounded-md bg-transparent p-1" />
                    )}
                </div>

                {/* --- Content Column --- */}
                <div className="flex-1 min-w-0 flex flex-col justify-center">
                    
                    {/* File Name */}
                    <div className={`text-sm font-semibold truncate pr-2 ${status === 'completed' ? 'text-emerald-900' : 'text-slate-700'}`}>
                        {name}
                    </div>

                    {/* Metadata Row */}
                    <div className="flex items-center justify-between mt-0.5 w-full">
                        <div className={`text-xs font-medium whitespace-nowrap truncate min-w-0 ${status === 'completed' ? 'text-emerald-600' : 'text-slate-400'}`}>
                            {sizeKB !== null ? `${sizeKB} KB` : '—'}
                            {rows !== undefined && (
                                <>
                                    <span className={`mx-1.5 ${status === 'completed' ? 'text-emerald-400' : 'text-slate-300'}`}>•</span>
                                    <span>{rows} rows</span>
                                </>
                            )}
                        </div>

                        {/* Badges */}
                        {(type || action) && (
                            <div className="flex items-center gap-1.5 ml-2 flex-none">
                                {type && (
                                    <span className={`inline-flex items-center px-2 py-0.5 rounded-full text-[9px] font-bold border uppercase tracking-wide whitespace-nowrap ${status === 'completed' ? 'bg-white text-emerald-600 border-emerald-100' : 'bg-indigo-50 text-indigo-600 border-indigo-100'}`}>
                                        {type}
                                    </span>
                                )}
                                {action && (
                                    <span className={`inline-flex items-center px-2 py-0.5 rounded-full text-[9px] font-bold border uppercase tracking-wide whitespace-nowrap ${status === 'completed' ? 'bg-white text-emerald-600 border-emerald-100' : 'bg-orange-50 text-orange-600 border-orange-100'}`}>
                                        {action}
                                    </span>
                                )}
                            </div>
                        )}
                    </div>
                </div>
            </div>
        </div>
    );
}