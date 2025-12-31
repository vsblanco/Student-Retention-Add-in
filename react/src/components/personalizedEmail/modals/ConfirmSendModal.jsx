import React from 'react';

export default function ConfirmSendModal({ isOpen, onClose, onConfirm, count }) {
    if (!isOpen) return null;

    return (
        <div className="fixed inset-0 bg-gray-600 bg-opacity-75 flex items-center justify-center z-50">
            <div className="bg-white rounded-lg shadow-xl p-6 w-full max-w-md">
                <h3 className="text-lg font-semibold text-gray-800 mb-4">Confirm Send</h3>
                <p className="text-sm text-gray-600">
                    You are about to send {count} {count === 1 ? 'email' : 'emails'}. Do you want to proceed?
                </p>
                <div className="flex justify-end gap-2 mt-4">
                    <button
                        onClick={onClose}
                        className="px-4 py-2 bg-gray-200 text-gray-700 rounded-md hover:bg-gray-300"
                    >
                        Cancel
                    </button>
                    <button
                        onClick={onConfirm}
                        className="px-4 py-2 bg-blue-600 text-white rounded-md hover:bg-blue-700"
                    >
                        Confirm & Send
                    </button>
                </div>
            </div>
        </div>
    );
}
