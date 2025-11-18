/* * Timestamp: 2025-11-18 17:05:00 EST
 * Version: 4.0.0
 * Author: Gemini (for Victor)
 * Description: FileCard with 'loading' state featuring skeleton text and a spinner ring.
 */

import React from 'react';
import csvIcon from '../../assets/icons/csv-icon.png';
import { File, CheckCircle, Loader2 } from 'lucide-react';

export default function FileCard({ file, rows, type, action, icon, status = 'normal' } = {}) {
    const name = (file && (file.name || file.filename)) || 'Unknown.csv';
    const sizeKB = file && file.size ? Math.round(file.size / 1024) : null;

    // 1. Define styles based on status
    const stateStyles = {
        normal: 'bg-white border-slate-100 shadow-sm hover:shadow-md hover:border-slate-200',
        completed: 'bg-emerald-50/60 border-emerald-200 shadow-sm',
        pending: 'bg-slate-50 border-slate-100 opacity-60 grayscale pointer-events-none',
        loading: 'bg-white border-slate-100 shadow-sm cursor-wait', // Loading base style
    };

    // 2. Define icon container styles
    const iconStyles = {
        normal: 'bg-slate-50 border-slate-100',
        completed: 'bg-white border-emerald-100',
        pending: 'bg-slate-100 border-slate-200',
        loading: 'bg-indigo-50/30 border-indigo-100', // Subtle tint for loading
    };

    return (
        <div
            title={name}
            className={`
                w-full group relative
                flex items-center gap-3
                p-3
                rounded-xl
                border
                transition-all duration-300 ease-in-out
                ${stateStyles[status] || stateStyles.normal}
            `}
        >
            {/* --- Icon Column --- */}
            <div className={`
                w-11 h-11 flex-none flex items-center justify-center 
                rounded-lg border overflow-hidden relative
                ${iconStyles[status] || iconStyles.normal}
            `}>
                
                {/* LOADING STATE: Fancy Spinner Loop */}
                {status === 'loading' && (
                    <>
                        {/* The spinning ring */}
                        <div className="absolute inset-0 rounded-lg border-2 border-indigo-600/20"></div>
                        <div className="absolute inset-0 rounded-lg border-2 border-indigo-500 border-t-transparent animate-spin"></div>
                        {/* Faded Icon behind spinner */}
                        <File size={20} className="text-indigo-300 opacity-50 scale-75" />
                    </>
                )}

                {/* COMPLETED STATE: Checkmark Overlay */}
                {status === 'completed' && (
                    <div className="absolute top-0 right-0 bg-emerald-500 w-3 h-3 rounded-bl-md flex items-center justify-center z-10">
                        <CheckCircle size={8} className="text-white" />
                    </div>
                )}

                {/* NORMAL / PENDING / COMPLETED Icons */}
                {status !== 'loading' && (
                    icon ? (
                        <img src={icon} alt="import type" className="w-8 h-8 object-contain rounded-md" />
                    ) : /\.csv$/i.test(name) ? (
                        <img src={csvIcon} alt="CSV" className="w-8 h-8 object-contain rounded-md" />
                    ) : (
                        <File size={20} className="text-slate-400 rounded-md bg-transparent p-1" />
                    )
                )}
            </div>

            {/* --- Content Column --- */}
            <div className="flex-1 min-w-0 flex flex-col justify-center">
                
                {/* LOADING STATE: Skeleton Text Bars */}
                {status === 'loading' ? (
                    <div className="space-y-2 w-full animate-pulse">
                        {/* Title Skeleton */}
                        <div className="h-3.5 bg-slate-200 rounded w-3/4"></div>
                        {/* Metadata Skeleton */}
                        <div className="flex items-center gap-2">
                            <div className="h-2.5 bg-slate-100 rounded w-1/4"></div>
                            <div className="h-2.5 bg-slate-100 rounded w-1/3"></div>
                        </div>
                    </div>
                ) : (
                    /* REAL CONTENT */
                    <>
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
                    </>
                )}
            </div>
        </div>
    );
}