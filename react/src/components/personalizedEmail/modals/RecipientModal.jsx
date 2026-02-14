import React, { useState, useEffect } from 'react';

export default function RecipientModal({
    isOpen,
    onClose,
    currentSelection,
    onConfirm,
    getStudentDataCore,
    recipientDataCache
}) {
    const [selection, setSelection] = useState(currentSelection);
    const [studentCount, setStudentCount] = useState(0);
    const [excludedStudents, setExcludedStudents] = useState([]);
    const [statusMessage, setStatusMessage] = useState('');
    const [isLoading, setIsLoading] = useState(false);
    const [showExclusionsView, setShowExclusionsView] = useState(false);

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
            const result = await getStudentDataCore(sel, true, [], (percent) => {
                setStatusMessage(`Counting students... ${percent}%`);
            });
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
            const result = await getStudentDataCore(sel, true, [], (percent) => {
                setStatusMessage(`Counting students... ${percent}%`);
            });
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
    };

    if (!isOpen) return null;

    const activeExclusionCount = (selection.excludeDNC ? 1 : 0) + (selection.excludeFillColor ? 1 : 0);

    return (
        <div className="fixed inset-0 bg-gray-600 bg-opacity-75 flex items-center justify-center z-50">
            <div className="bg-white rounded-lg shadow-xl p-6 w-full max-w-md">

                {showExclusionsView ? (
                    <>
                        <div className="flex items-center gap-2 mb-4">
                            <button
                                onClick={() => setShowExclusionsView(false)}
                                className="p-1 rounded-md hover:bg-gray-100 text-gray-500 hover:text-gray-700"
                                title="Back to Select Recipients"
                            >
                                <svg className="h-5 w-5" xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" strokeWidth="1.5" stroke="currentColor">
                                    <path strokeLinecap="round" strokeLinejoin="round" d="M15.75 19.5L8.25 12l7.5-7.5" />
                                </svg>
                            </button>
                            <h3 className="text-lg font-semibold text-gray-800">Exclusions</h3>
                        </div>

                        <div className="space-y-2">
                            <div className="bg-gray-50 p-3 rounded-md">
                                <label htmlFor="exclude-dnc-toggle" className="flex items-center justify-between cursor-pointer">
                                    <span className="text-sm text-gray-700 flex-grow pr-4">
                                        DNC-tagged students
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

                            <div className="bg-gray-50 p-3 rounded-md">
                                <label htmlFor="exclude-fill-color-toggle" className="flex items-center justify-between cursor-pointer">
                                    <span className="text-sm text-gray-700 flex-grow pr-4">
                                        Outreach Color Filled students
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

                        {excludedStudents.length > 0 && (
                            <div className="mt-4 border-t pt-3">
                                <p className="text-xs font-medium text-gray-500 mb-2">
                                    {excludedStudents.length} student{excludedStudents.length !== 1 ? 's' : ''} excluded
                                </p>
                                <div className="max-h-48 overflow-y-auto border border-gray-200 rounded-md">
                                    <ul className="divide-y divide-gray-200">
                                        {excludedStudents.map((student, index) => (
                                            <li key={index} className="px-3 py-2 flex justify-between items-center">
                                                <span className="text-sm text-gray-700 truncate pr-2" title={student.name}>
                                                    {student.name}
                                                </span>
                                                <span
                                                    className="flex-shrink-0 px-2 py-0.5 rounded-full text-xs"
                                                    style={
                                                        student.reason === 'DNC Tag'
                                                            ? { backgroundColor: '#FEE2E2', color: '#991B1B' }
                                                            : student.reason === 'Fill Color' && student.color
                                                                ? { backgroundColor: student.color, color: '#374151' }
                                                                : { backgroundColor: '#E5E7EB', color: '#4B5563' }
                                                    }
                                                >
                                                    {student.reason}
                                                </span>
                                            </li>
                                        ))}
                                    </ul>
                                </div>
                            </div>
                        )}

                        {excludedStudents.length === 0 && (selection.excludeDNC || selection.excludeFillColor) && (
                            <div className="mt-4 border-t pt-3">
                                <p className="text-xs text-gray-400 text-center">No students excluded</p>
                            </div>
                        )}
                    </>
                ) : (
                    <>
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
                            <button
                                onClick={() => setShowExclusionsView(true)}
                                className="w-full flex items-center justify-between px-3 py-2.5 bg-gray-50 rounded-md hover:bg-gray-100 transition-colors"
                            >
                                <span className="text-sm font-medium text-gray-700">Exclusions</span>
                                <div className="flex items-center gap-2">
                                    {activeExclusionCount > 0 && (
                                        <span className="px-2 py-0.5 text-xs font-medium bg-blue-100 text-blue-700 rounded-full">
                                            {activeExclusionCount} active
                                        </span>
                                    )}
                                    {excludedStudents.length > 0 && (
                                        <span className="px-2 py-0.5 text-xs font-medium bg-red-100 text-red-700 rounded-full">
                                            {excludedStudents.length} excluded
                                        </span>
                                    )}
                                    <svg className="h-4 w-4 text-gray-400" xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" strokeWidth="1.5" stroke="currentColor">
                                        <path strokeLinecap="round" strokeLinejoin="round" d="M8.25 4.5l7.5 7.5-7.5 7.5" />
                                    </svg>
                                </div>
                            </button>
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
                    </>
                )}

            </div>
        </div>
    );
}
