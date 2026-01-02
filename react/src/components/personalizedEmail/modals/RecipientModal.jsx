import React, { useState, useEffect } from 'react';

export default function RecipientModal({
    isOpen,
    onClose,
    currentSelection,
    onConfirm,
    getStudentDataCore,
    recipientDataCache,
    onDataFetch
}) {
    const [selection, setSelection] = useState(currentSelection);
    const [studentCount, setStudentCount] = useState(0);
    const [excludedStudents, setExcludedStudents] = useState([]);
    const [statusMessage, setStatusMessage] = useState('');
    const [isLoading, setIsLoading] = useState(false);
    const [showExcludedList, setShowExcludedList] = useState(false);

    useEffect(() => {
        if (isOpen) {
            setSelection(currentSelection);
            fetchStudentCount(currentSelection);
        }
    }, [isOpen, currentSelection]);

    const fetchStudentCount = async (sel = selection) => {
        setStatusMessage('Counting students...');
        setIsLoading(true);
        setExcludedStudents([]);

        const { type, excludeDNC, excludeFillColor } = sel;

        // Check cache first
        if (type !== 'custom' && excludeDNC && excludeFillColor && recipientDataCache.has(type)) {
            const cachedResult = recipientDataCache.get(type);
            setStudentCount(cachedResult.included.length);
            setExcludedStudents(cachedResult.excluded);
            setStatusMessage(`${cachedResult.included.length} student${cachedResult.included.length !== 1 ? 's' : ''} found.`);
            setIsLoading(false);
            return;
        }

        try {
            const result = await getStudentDataCore(sel, true); // Skip special params for faster counting
            setStudentCount(result.included.length);
            setExcludedStudents(result.excluded);
            setStatusMessage(`${result.included.length} student${result.included.length !== 1 ? 's' : ''} found.`);
            setIsLoading(false);
        } catch (error) {
            setStudentCount(0);
            setExcludedStudents([]);
            setStatusMessage(error.userFacingMessage || (error.userFacing ? error.message : 'An error occurred.'));
            setIsLoading(false);
        }
    };

    const handleSelectionChange = (field, value) => {
        const newSelection = { ...selection, [field]: value };
        setSelection(newSelection);
        // Re-fetch count when selection changes
        setTimeout(() => {
            fetchStudentCountForSelection(newSelection);
        }, 100);
    };

    const fetchStudentCountForSelection = async (sel) => {
        setStatusMessage('Counting students...');
        setIsLoading(true);
        setExcludedStudents([]);

        try {
            const result = await getStudentDataCore(sel, true); // Skip special params for faster counting
            setStudentCount(result.included.length);
            setExcludedStudents(result.excluded);
            setStatusMessage(`${result.included.length} student${result.included.length !== 1 ? 's' : ''} found.`);
            setIsLoading(false);
        } catch (error) {
            setStudentCount(0);
            setExcludedStudents([]);
            setStatusMessage(error.userFacingMessage || (error.userFacing ? error.message : 'An error occurred.'));
            setIsLoading(false);
        }
    };

    const handleConfirm = () => {
        onConfirm(selection, studentCount);
        onClose();
        onDataFetch().catch(() => {
            onConfirm(selection, -1);
        });
    };

    if (!isOpen) return null;

    return (
        <div className="fixed inset-0 bg-gray-600 bg-opacity-75 flex items-center justify-center z-50">
            <div className="bg-white rounded-lg shadow-xl p-6 w-full max-w-md">
                <h3 className="text-lg font-semibold text-gray-800 mb-4">Select Recipients</h3>

                <div className="space-y-4">
                    <p className="text-sm font-medium text-gray-700">Select students from:</p>
                    <fieldset className="space-y-2">
                        <div>
                            <input
                                type="radio"
                                id="recipient-lda"
                                name="recipient-source"
                                value="lda"
                                checked={selection.type === 'lda'}
                                onChange={(e) => handleSelectionChange('type', e.target.value)}
                                className="h-4 w-4 text-blue-600 border-gray-300 focus:ring-blue-500"
                            />
                            <label htmlFor="recipient-lda" className="ml-3 text-sm text-gray-700">
                                Today's LDA Sheet
                            </label>
                        </div>
                        <div>
                            <input
                                type="radio"
                                id="recipient-master"
                                name="recipient-source"
                                value="master"
                                checked={selection.type === 'master'}
                                onChange={(e) => handleSelectionChange('type', e.target.value)}
                                className="h-4 w-4 text-blue-600 border-gray-300 focus:ring-blue-500"
                            />
                            <label htmlFor="recipient-master" className="ml-3 text-sm text-gray-700">
                                Master List
                            </label>
                        </div>
                        <div>
                            <input
                                type="radio"
                                id="recipient-custom"
                                name="recipient-source"
                                value="custom"
                                checked={selection.type === 'custom'}
                                onChange={(e) => handleSelectionChange('type', e.target.value)}
                                className="h-4 w-4 text-blue-600 border-gray-300 focus:ring-blue-500"
                            />
                            <label htmlFor="recipient-custom" className="ml-3 text-sm text-gray-700">
                                Custom Sheet
                            </label>
                        </div>
                    </fieldset>

                    {selection.type === 'custom' && (
                        <div className="ml-7">
                            <label htmlFor="recipient-custom-sheet-name" className="block text-xs font-medium text-gray-600">
                                Custom Sheet Name
                            </label>
                            <input
                                type="text"
                                id="recipient-custom-sheet-name"
                                value={selection.customSheetName}
                                onChange={(e) => handleSelectionChange('customSheetName', e.target.value)}
                                className="mt-1 block w-full px-3 py-2 border border-gray-300 rounded-md shadow-sm text-sm"
                                placeholder="Enter the exact sheet name"
                            />
                        </div>
                    )}
                </div>

                <div className="mt-4 border-t pt-4">
                    <div className="flex items-center justify-between">
                        <p className="text-sm font-medium text-gray-700">Exclusions</p>
                        {excludedStudents.length > 0 && (
                            <div className="relative">
                                <button
                                    onClick={() => setShowExcludedList(!showExcludedList)}
                                    className="p-1 rounded-full hover:bg-gray-200"
                                    title={`${excludedStudents.length} student(s) excluded. Click to view.`}
                                >
                                    <svg className="h-5 w-5 text-gray-600" xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" strokeWidth="1.5" stroke="currentColor">
                                        <path strokeLinecap="round" strokeLinejoin="round" d="M8.25 6.75h7.5M8.25 12h7.5m-7.5 5.25h7.5M3.75 6.75h.007v.008H3.75V6.75zm.375 0a.375.375 0 11-.75 0 .375.375 0 01.75 0zM3.75 12h.007v.008H3.75V12zm.375 0a.375.375 0 11-.75 0 .375.375 0 01.75 0zm-.375 5.25h.007v.008H3.75v-.008zm.375 0a.375.375 0 11-.75 0 .375.375 0 01.75 0z" />
                                    </svg>
                                </button>
                                {showExcludedList && (
                                    <div className="absolute right-0 bottom-full mb-2 w-72 bg-white border border-gray-300 rounded-lg shadow-xl z-10">
                                        <div className="p-3 border-b">
                                            <h4 className="text-sm font-semibold text-gray-800">Excluded Students</h4>
                                        </div>
                                        <div className="max-h-48 overflow-y-auto p-2 text-xs">
                                            <ul className="divide-y divide-gray-200">
                                                {excludedStudents.map((student, index) => (
                                                    <li key={index} className="p-2 flex justify-between items-center">
                                                        <span className="text-gray-700 truncate pr-2" title={student.name}>
                                                            {student.name}
                                                        </span>
                                                        <span className="flex-shrink-0 px-2 py-0.5 bg-gray-200 text-gray-600 rounded-full text-xs">
                                                            {student.reason}
                                                        </span>
                                                    </li>
                                                ))}
                                            </ul>
                                        </div>
                                    </div>
                                )}
                            </div>
                        )}
                    </div>

                    <div className="bg-gray-50 p-3 rounded-md">
                        <label htmlFor="exclude-dnc-toggle" className="flex items-center justify-between cursor-pointer">
                            <span className="text-sm text-gray-700 flex-grow pr-4">
                                Exclude students with a "DNC" tag
                            </span>
                            <div className="relative inline-flex items-center flex-shrink-0">
                                <input
                                    type="checkbox"
                                    id="exclude-dnc-toggle"
                                    checked={selection.excludeDNC}
                                    onChange={(e) => handleSelectionChange('excludeDNC', e.target.checked)}
                                    className="sr-only peer"
                                />
                                <div className="w-11 h-6 bg-gray-200 rounded-full peer peer-focus:ring-4 peer-focus:ring-blue-300 peer-checked:after:translate-x-full peer-checked:after:border-white after:content-[''] after:absolute after:top-0.5 after:left-[2px] after:bg-white after:border-gray-300 after:border after:rounded-full after:h-5 after:w-5 after:transition-all peer-checked:bg-blue-600"></div>
                            </div>
                        </label>
                    </div>

                    <div className="mt-2 bg-gray-50 p-3 rounded-md">
                        <label htmlFor="exclude-fill-color-toggle" className="flex items-center justify-between cursor-pointer">
                            <span className="text-sm text-gray-700 flex-grow pr-4">
                                Exclude students with a Fill Color in their Outreach column
                            </span>
                            <div className="relative inline-flex items-center flex-shrink-0">
                                <input
                                    type="checkbox"
                                    id="exclude-fill-color-toggle"
                                    checked={selection.excludeFillColor}
                                    onChange={(e) => handleSelectionChange('excludeFillColor', e.target.checked)}
                                    className="sr-only peer"
                                />
                                <div className="w-11 h-6 bg-gray-200 rounded-full peer peer-focus:ring-4 peer-focus:ring-blue-300 peer-checked:after:translate-x-full peer-checked:after:border-white after:content-[''] after:absolute after:top-0.5 after:left-[2px] after:bg-white after:border-gray-300 after:border after:rounded-full after:h-5 after:w-5 after:transition-all peer-checked:bg-blue-600"></div>
                            </div>
                        </label>
                    </div>
                </div>

                <div className="mt-6 h-8 flex items-center justify-center">
                    <p className="text-sm text-gray-600">{statusMessage}</p>
                </div>

                <div className="flex justify-end gap-2 mt-4 border-t pt-4">
                    <button
                        onClick={onClose}
                        className="px-4 py-2 bg-gray-200 text-gray-700 rounded-md hover:bg-gray-300"
                    >
                        Cancel
                    </button>
                    <button
                        onClick={handleConfirm}
                        disabled={isLoading}
                        className="px-4 py-2 bg-blue-600 text-white rounded-md hover:bg-blue-700 disabled:opacity-50 disabled:cursor-not-allowed"
                    >
                        Confirm
                    </button>
                </div>
            </div>
        </div>
    );
}
