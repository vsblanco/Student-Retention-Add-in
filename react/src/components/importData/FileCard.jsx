/* * Timestamp: 2025-11-22 14:17:00 EST
 * Version: 1.3.0 (Demo)
 * Author: Gemini (for Victor)
 * Description: Updated loading state to use a 3-dot scaling animation instead of a spinner.
 */

import React, { useState } from 'react';
import { File, CheckCircle, Loader2, RefreshCw, Shield, AlertCircle, Zap } from 'lucide-react';

// --- ASSET MOCK (Since we don't have your local files) ---
const csvIcon = `data:image/svg+xml;utf8,<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 32 32" fill="none" stroke="%23f97316" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect x="4" y="4" width="24" height="24" rx="4" fill="%23fff7ed"/><path d="M10 10h12"/><path d="M10 16h12"/><path d="M10 22h8"/></svg>`;

// --- YOUR COMPONENT (FileCard.jsx) ---

function FileCard({ file, rows, type, action, icon, status = 'normal' } = {}) {
    const name = (file && (file.name || file.filename)) || 'Unknown.csv';
    const sizeKB = file && file.size ? Math.round(file.size / 1024) : null;

    // 1. Define styles based on status
    const stateStyles = {
        normal: 'bg-white border-slate-100 shadow-sm hover:shadow-md hover:border-slate-200',
        completed: 'bg-emerald-50/60 border-emerald-200 shadow-sm',
        pending: 'bg-slate-50 border-slate-100 opacity-60 grayscale pointer-events-none',
        // Loading: Looks like pending (bg-slate-50), but no opacity on ROOT so spinner pops
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
            {/* Inject custom keyframes for the dot animation */}
            <style>
                {`
                    @keyframes dot-scale {
                        0%, 100% { transform: scale(0.75); opacity: 0.5; }
                        50% { transform: scale(1.5); opacity: 1; }
                    }
                `}
            </style>

            {/* --- NEW LOADING OVERLAY (3 Dots) --- */}
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
            {/* We fade this out if loading to mimic 'pending', but keep spinner bright */}
            <div className={`flex items-center gap-3 ${status === 'loading' ? 'opacity-50 grayscale blur-[0.5px]' : ''}`}>
                
                {/* --- Icon Column --- */}
                <div className={`
                    w-11 h-11 flex-none flex items-center justify-center 
                    rounded-lg border overflow-hidden relative
                    ${iconStyles[status] || iconStyles.normal}
                `}>
                    
                    {/* COMPLETED STATE: Checkmark Removed per request v1.2.0 */}
                    
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


// --- DEMO APP ---

export default function App() {
    const [activeStatus, setActiveStatus] = useState('normal');
    const [fileName, setFileName] = useState('customer_data_2024.csv');
    
    // Files for the gallery
    const demoFiles = [
        {
            title: "Normal State",
            description: "Default view with metadata.",
            props: {
                status: "normal",
                file: { name: "leads_q3_2024.csv", size: 4500000 },
                rows: 15420,
                type: "Import",
                action: "Process"
            }
        },
        {
            title: "Loading State",
            description: "Pending-style with 3 pulsing dots.",
            props: {
                status: "loading",
                file: { name: "large_dataset_upload.csv", size: 2400000 },
                rows: 8200,
                type: "Uploading"
            }
        },
        {
            title: "Completed State",
            description: "Emerald theme.",
            props: {
                status: "completed",
                file: { name: "clean_export_final.csv", size: 1200000 },
                rows: 8500,
                action: "Download"
            }
        },
        {
            title: "Pending State",
            description: "Greyscale & reduced opacity.",
            props: {
                status: "pending",
                file: { name: "waiting_in_queue.csv", size: 1024 },
                type: "Queue"
            }
        }
    ];

    return (
        <div className="min-h-screen bg-slate-50 p-6 md:p-12 font-sans text-slate-800">
            <div className="max-w-4xl mx-auto space-y-12">
                
                {/* Header */}
                <div className="text-center space-y-2">
                    <h1 className="text-3xl font-bold text-slate-900">FileCard Component Demo</h1>
                    <p className="text-slate-500">Visualizing the different lifecycle states of your file component.</p>
                </div>

                {/* 1. Interactive Playground */}
                <section className="bg-white rounded-2xl border border-slate-200 shadow-sm p-6 md:p-8">
                    <div className="flex items-center gap-2 mb-6 pb-4 border-b border-slate-100">
                        <Zap className="text-indigo-500" size={20} />
                        <h2 className="text-lg font-semibold">Interactive Playground</h2>
                    </div>

                    <div className="grid grid-cols-1 md:grid-cols-2 gap-8 lg:gap-12 items-center">
                        {/* Controls */}
                        <div className="space-y-6">
                            <div className="space-y-3">
                                <label className="text-xs font-bold text-slate-400 uppercase tracking-wider">Set Status</label>
                                <div className="grid grid-cols-2 gap-2">
                                    {['normal', 'loading', 'completed', 'pending'].map((s) => (
                                        <button
                                            key={s}
                                            onClick={() => setActiveStatus(s)}
                                            className={`
                                                px-3 py-2 rounded-lg text-sm font-medium capitalize transition-all
                                                ${activeStatus === s 
                                                    ? 'bg-indigo-600 text-white shadow-md shadow-indigo-200 ring-2 ring-indigo-100' 
                                                    : 'bg-slate-50 text-slate-600 hover:bg-slate-100'
                                                }
                                            `}
                                        >
                                            {s}
                                        </button>
                                    ))}
                                </div>
                            </div>
                            
                            <div className="space-y-3">
                                <label className="text-xs font-bold text-slate-400 uppercase tracking-wider">File Name</label>
                                <input 
                                    type="text" 
                                    value={fileName}
                                    onChange={(e) => setFileName(e.target.value)}
                                    className="w-full px-3 py-2 rounded-lg border border-slate-200 text-sm focus:outline-none focus:ring-2 focus:ring-indigo-500/20 focus:border-indigo-500"
                                />
                            </div>
                        </div>

                        {/* Preview Area */}
                        <div className="bg-slate-50/50 rounded-xl p-6 border border-dashed border-slate-200 flex flex-col items-center justify-center min-h-[200px]">
                            <div className="w-full max-w-sm transform transition-all duration-500 hover:scale-105">
                                <FileCard 
                                    status={activeStatus}
                                    file={{ name: fileName, size: 1024 * 500 }}
                                    rows={activeStatus === 'loading' ? undefined : 1234}
                                    type={activeStatus === 'loading' ? null : "Import"}
                                    action={activeStatus === 'loading' ? null : "Review"}
                                />
                            </div>
                            <p className="mt-4 text-xs text-slate-400 font-mono">
                                &lt;FileCard status="{activeStatus}" ... /&gt;
                            </p>
                        </div>
                    </div>
                </section>

                {/* 2. Static Gallery */}
                <section>
                    <div className="flex items-center gap-2 mb-6">
                        <Shield className="text-emerald-500" size={20} />
                        <h2 className="text-lg font-semibold text-slate-900">State Gallery</h2>
                    </div>

                    <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                        {demoFiles.map((demo, idx) => (
                            <div key={idx} className="bg-white p-4 rounded-xl border border-slate-200 shadow-sm flex flex-col gap-4">
                                <div className="pb-2 border-b border-slate-50">
                                    <h3 className="text-sm font-bold text-slate-700">{demo.title}</h3>
                                    <p className="text-xs text-slate-400">{demo.description}</p>
                                </div>
                                <div>
                                    <FileCard {...demo.props} />
                                </div>
                            </div>
                        ))}
                    </div>
                </section>

                {/* 3. Edge Cases */}
                <section>
                    <div className="flex items-center gap-2 mb-6">
                        <AlertCircle className="text-orange-500" size={20} />
                        <h2 className="text-lg font-semibold text-slate-900">Edge Cases</h2>
                    </div>
                    
                    <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
                        <FileCard file={{ name: "very_long_filename_that_should_truncate_correctly_in_the_ui_layout_test.csv", size: 2048 }} />
                        <FileCard file={{ name: "unknown_file_type.xyz", size: 2048 }} />
                        <FileCard file={{ name: "no_metadata.csv" }} />
                    </div>
                </section>

            </div>
        </div>
    );
}