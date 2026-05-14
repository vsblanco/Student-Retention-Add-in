import React, { useState } from 'react';

export default function MailMergeInstructionsModal({ isOpen, onClose, templateFilename, recipientsFilename, subject, count }) {
    const [copied, setCopied] = useState(false);

    if (!isOpen) return null;

    const handleCopySubject = async () => {
        try {
            await navigator.clipboard.writeText(subject || '');
            setCopied(true);
            setTimeout(() => setCopied(false), 1500);
        } catch (err) {
            console.error('Clipboard write failed:', err);
        }
    };

    return (
        <div className="fixed inset-0 bg-gray-600 bg-opacity-75 flex items-center justify-center z-50 p-4">
            <div className="bg-white rounded-lg shadow-xl p-6 w-full max-w-lg max-h-[90vh] overflow-y-auto">
                <h3 className="text-lg font-semibold text-gray-800 mb-2">Downloaded — now finish the merge in Word</h3>
                <p className="text-sm text-gray-600 mb-4">
                    Two files were downloaded for {count} {count === 1 ? 'recipient' : 'recipients'}:
                </p>

                <ul className="text-sm text-gray-700 mb-4 space-y-1">
                    <li className="flex items-baseline gap-2">
                        <span className="inline-block w-2 h-2 rounded-full bg-blue-500 flex-shrink-0" />
                        <code className="text-xs bg-gray-100 px-1.5 py-0.5 rounded">{templateFilename}</code>
                        <span className="text-gray-500">(the email body template)</span>
                    </li>
                    <li className="flex items-baseline gap-2">
                        <span className="inline-block w-2 h-2 rounded-full bg-green-500 flex-shrink-0" />
                        <code className="text-xs bg-gray-100 px-1.5 py-0.5 rounded">{recipientsFilename}</code>
                        <span className="text-gray-500">(recipient list)</span>
                    </li>
                </ul>

                <ol className="text-sm text-gray-700 space-y-2 mb-4 list-decimal pl-5">
                    <li>Open the <strong>template</strong> in Word.</li>
                    <li>Go to the <strong>Mailings</strong> tab → <strong>Select Recipients</strong> → <strong>Use an Existing List</strong> → pick the recipients <code className="text-xs bg-gray-100 px-1 rounded">.xlsx</code>.</li>
                    <li>Click <strong>Finish &amp; Merge</strong> → <strong>Send Email Messages…</strong></li>
                    <li>In the dialog: set <em>To</em> = <code className="text-xs bg-gray-100 px-1 rounded">Email</code>, paste the subject below, set Mail format = <strong>HTML</strong>, click OK.</li>
                </ol>

                {subject && (
                    <div className="mb-4">
                        <label className="block text-xs font-medium text-gray-600 mb-1">Subject (copy into Word's dialog)</label>
                        <div className="flex gap-2">
                            <input
                                type="text"
                                readOnly
                                value={subject}
                                className="flex-1 px-2 py-1 text-sm border border-gray-300 rounded bg-gray-50 text-gray-800"
                                onFocus={(e) => e.target.select()}
                            />
                            <button
                                onClick={handleCopySubject}
                                className="px-3 py-1 bg-blue-600 text-white text-sm rounded hover:bg-blue-700"
                            >
                                {copied ? 'Copied!' : 'Copy'}
                            </button>
                        </div>
                        <p className="text-xs text-gray-500 mt-1">
                            Word's Send Email dialog only takes plain text for the subject, so any <code>{'{Field}'}</code> placeholders in the subject won't be substituted.
                        </p>
                    </div>
                )}

                <div className="flex justify-end">
                    <button
                        onClick={onClose}
                        className="px-4 py-2 bg-gray-200 text-gray-700 rounded-md hover:bg-gray-300"
                    >
                        Got it
                    </button>
                </div>
            </div>
        </div>
    );
}
