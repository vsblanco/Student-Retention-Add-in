import React, { useState } from 'react';

export default function DownloadModal({
    isOpen,
    onClose,
    onDownloadTemplate,
    onDownloadRecipients,
    recipientCount,
}) {
    const [downloading, setDownloading] = useState(null); // 'template' | 'recipients' | null
    const [doneWith, setDoneWith] = useState(null);

    if (!isOpen) return null;

    const trigger = async (kind, fn) => {
        if (downloading) return;
        setDownloading(kind);
        try {
            await fn();
            setDoneWith(kind);
            setTimeout(() => setDoneWith(curr => (curr === kind ? null : curr)), 1500);
        } finally {
            setDownloading(null);
        }
    };

    const tile = ({ kind, label, sublabel, disabled, color, icon, onClick }) => {
        const isBusy = downloading === kind;
        const isDone = doneWith === kind;
        return (
            <button
                type="button"
                onClick={onClick}
                disabled={disabled || !!downloading}
                className={`flex flex-col items-center justify-center gap-2 p-4 border-2 rounded-lg transition-all w-full ${
                    disabled
                        ? 'border-gray-200 bg-gray-50 opacity-50 cursor-not-allowed'
                        : isDone
                            ? 'border-green-400 bg-green-50'
                            : `border-gray-200 ${color.hoverBorder} hover:bg-gray-50`
                }`}
            >
                <div className={`flex items-center justify-center h-12 w-12 rounded ${color.bg}`}>
                    {icon}
                </div>
                <div className="text-sm font-semibold text-gray-800">{label}</div>
                <div className="text-xs text-gray-500 h-4">
                    {isBusy ? 'Generating…' : isDone ? 'Downloaded' : sublabel}
                </div>
            </button>
        );
    };

    // Stylized doc-with-letter glyph in product brand colors.
    const WordIcon = (
        <svg viewBox="0 0 24 24" className="h-7 w-7" fill="white">
            <path d="M14 2H6a2 2 0 00-2 2v16a2 2 0 002 2h12a2 2 0 002-2V8l-6-6z" />
            <path d="M14 2v6h6" fill="#2B579A" />
            <text x="12" y="17" fontSize="6" fontWeight="bold" fill="#2B579A" textAnchor="middle" fontFamily="Arial">W</text>
        </svg>
    );
    const ExcelIcon = (
        <svg viewBox="0 0 24 24" className="h-7 w-7" fill="white">
            <path d="M14 2H6a2 2 0 00-2 2v16a2 2 0 002 2h12a2 2 0 002-2V8l-6-6z" />
            <path d="M14 2v6h6" fill="#217346" />
            <text x="12" y="17" fontSize="6" fontWeight="bold" fill="#217346" textAnchor="middle" fontFamily="Arial">X</text>
        </svg>
    );

    return (
        <div className="fixed inset-0 bg-gray-600 bg-opacity-75 flex items-center justify-center z-50 p-4">
            <div className="bg-white rounded-lg shadow-xl p-6 w-full max-w-md">
                <div className="flex justify-between items-center mb-2">
                    <h3 className="text-lg font-semibold text-gray-800">Download Mail Merge Files</h3>
                    <button
                        onClick={onClose}
                        className="text-gray-400 hover:text-gray-600"
                        aria-label="Close"
                    >
                        <svg className="h-6 w-6" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                            <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M6 18L18 6M6 6l12 12" />
                        </svg>
                    </button>
                </div>
                <p className="text-xs text-gray-500 mb-4">
                    Grab one or both. If you&apos;re reusing the same template you can skip re-downloading it.
                </p>

                <div className="grid grid-cols-2 gap-3">
                    {tile({
                        kind: 'template',
                        label: 'Template',
                        sublabel: 'Word .docx',
                        disabled: false,
                        color: { bg: 'bg-[#2B579A]', hoverBorder: 'hover:border-blue-500' },
                        icon: WordIcon,
                        onClick: () => trigger('template', onDownloadTemplate),
                    })}
                    {tile({
                        kind: 'recipients',
                        label: 'Recipient List',
                        sublabel: `Excel .xlsx${recipientCount ? ` (${recipientCount})` : ''}`,
                        disabled: !recipientCount,
                        color: { bg: 'bg-[#217346]', hoverBorder: 'hover:border-green-500' },
                        icon: ExcelIcon,
                        onClick: () => trigger('recipients', onDownloadRecipients),
                    })}
                </div>

                <div className="flex justify-end mt-4">
                    <button
                        onClick={onClose}
                        className="px-4 py-2 bg-gray-200 text-gray-700 rounded-md hover:bg-gray-300 text-sm"
                    >
                        Done
                    </button>
                </div>
            </div>
        </div>
    );
}
