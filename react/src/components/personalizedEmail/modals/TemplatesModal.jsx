import React, { useState, useEffect } from 'react';
import { EMAIL_TEMPLATES_KEY } from '../utils/constants';

export default function TemplatesModal({ isOpen, onClose, onLoadTemplate, user, currentFrom, currentSubject, currentBody, currentCC }) {
    const [templates, setTemplates] = useState([]);
    const [expandedAuthors, setExpandedAuthors] = useState(new Set());
    const [showSaveModal, setShowSaveModal] = useState(false);
    const [editingTemplate, setEditingTemplate] = useState(null);
    const [templateName, setTemplateName] = useState('');
    const [saveStatus, setSaveStatus] = useState('');

    const isGuest = user === 'Guest';

    useEffect(() => {
        if (isOpen) {
            loadTemplates();
        }
    }, [isOpen]);

    const loadTemplates = async () => {
        const loaded = await getTemplates();
        setTemplates(loaded);
    };

    const getTemplates = async () => {
        return Excel.run(async (context) => {
            const settings = context.workbook.settings;
            const templatesSetting = settings.getItemOrNullObject(EMAIL_TEMPLATES_KEY);
            templatesSetting.load("value");
            await context.sync();
            return templatesSetting.value ? JSON.parse(templatesSetting.value) : [];
        });
    };

    const saveTemplates = async (templatesArray) => {
        await Excel.run(async (context) => {
            context.workbook.settings.add(EMAIL_TEMPLATES_KEY, JSON.stringify(templatesArray));
            await context.sync();
        });
        setTemplates(templatesArray);
    };

    const toggleAuthor = (author) => {
        const newExpanded = new Set(expandedAuthors);
        if (newExpanded.has(author)) {
            newExpanded.delete(author);
        } else {
            newExpanded.add(author);
        }
        setExpandedAuthors(newExpanded);
    };

    const handleLoadTemplate = (template) => {
        onLoadTemplate(template);
        onClose();
    };

    const handleEditTemplate = (template) => {
        setEditingTemplate(template);
        setTemplateName(template.name);
        setShowSaveModal(true);
    };

    const handleDeleteTemplate = async (template) => {
        const updatedTemplates = templates.filter(t => t.id !== template.id);
        await saveTemplates(updatedTemplates);
        setShowSaveModal(false);
    };

    const handleSaveTemplate = async () => {
        if (!templateName.trim()) {
            setSaveStatus('Template name is required.');
            return;
        }

        const templateData = {
            name: templateName.trim(),
            author: user || 'Unknown',
            from: currentFrom,
            subject: currentSubject,
            body: currentBody,
            cc: currentCC || []
        };

        let updatedTemplates = [...templates];
        if (editingTemplate) {
            const index = updatedTemplates.findIndex(t => t.id === editingTemplate.id);
            if (index !== -1) {
                updatedTemplates[index] = { ...updatedTemplates[index], ...templateData };
            }
        } else {
            updatedTemplates.push({
                ...templateData,
                id: 'tpl-' + Date.now(),
                createdAt: new Date().toISOString()
            });
        }

        await saveTemplates(updatedTemplates);
        setSaveStatus('Template saved successfully!');
        setTimeout(() => {
            setShowSaveModal(false);
            setSaveStatus('');
        }, 1500);
    };

    if (!isOpen) return null;

    // Group templates by author
    const groupedTemplates = templates.reduce((acc, template) => {
        const author = template.author || 'Uncategorized';
        if (!acc[author]) acc[author] = [];
        acc[author].push(template);
        return acc;
    }, {});

    return (
        <>
            <div className="fixed inset-0 bg-gray-600 bg-opacity-75 flex items-center justify-center z-50">
                <div className="bg-white rounded-lg shadow-xl p-6 w-full max-w-xl">
                    <h3 className="text-lg font-semibold text-gray-800 mb-4">Email Templates</h3>

                    <div className="space-y-2 max-h-80 overflow-y-auto border-t border-b py-4">
                        {templates.length === 0 ? (
                            <p className="text-center text-gray-500 text-sm">No saved templates yet.</p>
                        ) : (
                            Object.keys(groupedTemplates).sort().map(author => (
                                <div key={author}>
                                    <div
                                        onClick={() => toggleAuthor(author)}
                                        className="flex items-center justify-between p-2 rounded-md hover:bg-gray-100 cursor-pointer"
                                    >
                                        <div className="flex items-center space-x-2">
                                            <svg className="h-5 w-5 text-gray-500" viewBox="0 0 20 20" fill="currentColor">
                                                <path d="M2 6a2 2 0 012-2h5l2 2h5a2 2 0 012 2v6a2 2 0 01-2 2H4a2 2 0 01-2-2V6z"></path>
                                            </svg>
                                            <span className="font-semibold text-gray-700">{author}</span>
                                        </div>
                                        <svg
                                            className={`h-5 w-5 text-gray-500 transition-transform ${
                                                expandedAuthors.has(author) ? 'rotate-90' : ''
                                            }`}
                                            viewBox="0 0 20 20"
                                            fill="currentColor"
                                        >
                                            <path d="M7.293 14.707a1 1 0 010-1.414L10.586 10 7.293 6.707a1 1 0 011.414-1.414l4 4a1 1 0 010 1.414l-4 4a1 1 0 01-1.414 0z"></path>
                                        </svg>
                                    </div>

                                    {expandedAuthors.has(author) && (
                                        <div className="pl-6 border-l-2 border-gray-200 ml-2">
                                            {groupedTemplates[author].map(template => (
                                                <div
                                                    key={template.id}
                                                    className="flex items-center justify-between p-2 my-1 rounded-md hover:bg-gray-50"
                                                >
                                                    <div className="text-sm font-medium text-gray-800">{template.name}</div>
                                                    <div className="flex space-x-2">
                                                        <button
                                                            onClick={() => handleLoadTemplate(template)}
                                                            className="px-2 py-1 bg-blue-100 text-blue-800 text-xs font-semibold rounded-md hover:bg-blue-200"
                                                        >
                                                            Load
                                                        </button>
                                                        {!isGuest && (
                                                            <button
                                                                onClick={() => handleEditTemplate(template)}
                                                                className="px-2 py-1 bg-gray-200 text-gray-800 text-xs font-semibold rounded-md hover:bg-gray-300"
                                                            >
                                                                Edit
                                                            </button>
                                                        )}
                                                    </div>
                                                </div>
                                            ))}
                                        </div>
                                    )}
                                </div>
                            ))
                        )}
                    </div>

                    <div className={`flex ${isGuest ? 'justify-end' : 'justify-between'} mt-4`}>
                        {!isGuest && (
                            <button
                                onClick={() => {
                                    setEditingTemplate(null);
                                    setTemplateName('');
                                    setShowSaveModal(true);
                                }}
                                className="px-4 py-2 bg-green-600 text-white rounded-md hover:bg-green-700"
                            >
                                Save Current as Template
                            </button>
                        )}
                        <button
                            onClick={onClose}
                            className="px-4 py-2 bg-gray-200 text-gray-700 rounded-md hover:bg-gray-300"
                        >
                            Close
                        </button>
                    </div>
                </div>
            </div>

            {/* Save Template Modal */}
            {showSaveModal && (
                <div className="fixed inset-0 bg-gray-600 bg-opacity-75 flex items-center justify-center z-60">
                    <div className="bg-white rounded-lg shadow-xl p-6 w-full max-w-md">
                        <h3 className="text-lg font-semibold text-gray-800 mb-4">
                            {editingTemplate ? 'Edit Template' : 'Save Template'}
                        </h3>
                        <div className="space-y-4">
                            <div>
                                <label htmlFor="template-name" className="block text-sm font-medium text-gray-700">
                                    Template Name
                                </label>
                                <input
                                    type="text"
                                    id="template-name"
                                    value={templateName}
                                    onChange={(e) => setTemplateName(e.target.value)}
                                    className="mt-1 block w-full px-3 py-2 border border-gray-300 rounded-md shadow-sm"
                                    placeholder="e.g., Mid-term Check-in"
                                />
                            </div>
                            <div>
                                <label className="block text-sm font-medium text-gray-700">
                                    Author
                                </label>
                                <div className="mt-1 block w-full px-3 py-2 border border-gray-300 rounded-md shadow-sm bg-gray-100 text-gray-700">
                                    {user || 'Unknown'}
                                </div>
                            </div>
                        </div>
                        <p className="text-xs mt-2 h-4 text-center text-red-600">{saveStatus}</p>
                        <div className="flex justify-between items-center mt-4">
                            {editingTemplate && (
                                <button
                                    onClick={() => handleDeleteTemplate(editingTemplate)}
                                    className="px-4 py-2 bg-red-600 text-white rounded-md hover:bg-red-700"
                                >
                                    Delete
                                </button>
                            )}
                            <div className="flex gap-2 ml-auto">
                                <button
                                    onClick={() => {
                                        setShowSaveModal(false);
                                        setSaveStatus('');
                                    }}
                                    className="px-4 py-2 bg-gray-200 text-gray-700 rounded-md hover:bg-gray-300"
                                >
                                    Cancel
                                </button>
                                <button
                                    onClick={handleSaveTemplate}
                                    className="px-4 py-2 bg-blue-600 text-white rounded-md hover:bg-blue-700"
                                >
                                    Save
                                </button>
                            </div>
                        </div>
                    </div>
                </div>
            )}
        </>
    );
}
