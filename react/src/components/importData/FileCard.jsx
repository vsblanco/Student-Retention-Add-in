/* * Timestamp: 2025-11-18 15:55:00 EST
 * Version: 2.1.0
 * Author: Gemini (for Victor)
 * Description: Refactored FileCard - Fixed layout to prevent text wrapping.
 */

import React from 'react';
import csvIcon from '../../assets/icons/csv-icon.png';
import { File } from 'lucide-react';

export default function FileCard({ file, rows, type, action, icon } = {}) {
    const name = (file && (file.name || file.filename)) || 'Unknown.csv';
    const sizeKB = file && file.size ? Math.round(file.size / 1024) : null;

    return (
        <div
            title={name}
            className="
                w-full group relative
                flex items-center gap-3
                p-3
                bg-white rounded-xl
                border border-slate-100
                shadow-sm hover:shadow-md hover:border-slate-200
                transition-all duration-200 ease-in-out
            "
        >
            {/* Icon Container (Fixed Width) */}
            <div className="w-11 h-11 flex-none flex items-center justify-center bg-slate-50 rounded-lg border border-slate-100">
                {icon ? (
                    <img src={icon} alt="import type" className="w-8 h-8 object-contain" />
                ) : /\.csv$/i.test(name) ? (
                    <img src={csvIcon} alt="CSV" className="w-8 h-8 object-contain" />
                ) : (
                    <File size={24} className="text-slate-400" />
                )}
            </div>

            {/* Content Container */}
            <div className="flex-1 min-w-0 flex flex-col justify-center">
                
                {/* File Name (Truncates if too long) */}
                <div className="text-sm font-semibold text-slate-700 truncate pr-2">
                    {name}
                </div>

                {/* Metadata Row (Flex-nowrap to prevent wrapping) */}
                <div className="flex items-center justify-between mt-0.5 w-full">
                    
                    {/* Text Info: Size & Rows (Nowrap + Truncate allowed) */}
                    <div className="text-xs text-slate-400 font-medium whitespace-nowrap truncate min-w-0">
                        {sizeKB !== null ? `${sizeKB} KB` : '—'}
                        {rows !== undefined && (
                            <>
                                <span className="mx-1.5 text-slate-300">•</span>
                                <span>{rows} rows</span>
                            </>
                        )}
                    </div>

                    {/* Badges Area (Fixed width, does not shrink) */}
                    {(type || action) && (
                        <div className="flex items-center gap-1.5 ml-2 flex-none">
                            {type && (
                                <span className="inline-flex items-center px-2 py-0.5 rounded-full text-[9px] font-bold bg-indigo-50 text-indigo-600 border border-indigo-100 uppercase tracking-wide whitespace-nowrap">
                                    {type}
                                </span>
                            )}
                            {/* Re-added Action Badge */}
                            {action && (
                                <span className="inline-flex items-center px-2 py-0.5 rounded-full text-[9px] font-bold bg-orange-50 text-gray-600 border border-gray-200 uppercase tracking-wide whitespace-nowrap">
                                    {action}
                                </span>
                            )}
                        </div>
                    )}
                </div>
            </div>
        </div>
    );
}