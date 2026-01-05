import React, { useState, useEffect } from 'react';

export default function PowerAutomateConfigModal({ isOpen, onClose }) {
    const [url, setUrl] = useState('');
    const [status, setStatus] = useState('');
    const [currentConnection, setCurrentConnection] = useState(null);
    const [isLoading, setIsLoading] = useState(false);

    useEffect(() => {
        if (isOpen) {
            loadCurrentConnection();
        }
    }, [isOpen]);

    const loadCurrentConnection = async () => {
        setIsLoading(true);
        try {
            await Excel.run(async (context) => {
                const settings = context.workbook.settings;
                const connectionsSetting = settings.getItemOrNullObject("connections");
                connectionsSetting.load("value");
                await context.sync();

                const connections = connectionsSetting.value ? JSON.parse(connectionsSetting.value) : [];
                const connection = connections.find(c => c.type === 'power-automate' && c.name === 'Send Personalized Email');

                if (connection) {
                    setCurrentConnection(connection);
                    setUrl(connection.url || '');
                    setStatus('Connection configured');
                } else {
                    setCurrentConnection(null);
                    setUrl('');
                    setStatus('No connection configured');
                }
            });
        } catch (error) {
            console.error('Error loading connection:', error);
            setStatus('Error loading connection');
        } finally {
            setIsLoading(false);
        }
    };

    const isValidHttpUrl = (string) => {
        try {
            const url = new URL(string);
            return url.protocol === "http:" || url.protocol === "https:";
        } catch {
            return false;
        }
    };

    const handleSave = async () => {
        if (!isValidHttpUrl(url)) {
            setStatus('Please enter a valid HTTP URL.');
            return;
        }

        setStatus('Saving connection...');

        try {
            await Excel.run(async (context) => {
                const settings = context.workbook.settings;
                const connectionsSetting = settings.getItemOrNullObject("connections");
                connectionsSetting.load("value");
                await context.sync();

                let connections = connectionsSetting.value ? JSON.parse(connectionsSetting.value) : [];

                // Remove existing connection if any
                connections = connections.filter(c => !(c.type === 'power-automate' && c.name === 'Send Personalized Email'));

                // Add new connection
                const newConnection = {
                    id: currentConnection?.id || ('pa-' + Math.random().toString(36).substr(2, 9)),
                    name: 'Send Personalized Email',
                    type: 'power-automate',
                    url: url,
                    actions: [],
                    history: []
                };
                connections.push(newConnection);

                settings.add("connections", JSON.stringify(connections));
                await context.sync();

                setCurrentConnection(newConnection);
                setStatus('Connection saved successfully!');
                setTimeout(() => {
                    onClose();
                }, 1500);
            });
        } catch (error) {
            console.error('Error saving connection:', error);
            setStatus('Error saving connection');
        }
    };

    const handleDelete = async () => {
        if (!currentConnection) return;

        setStatus('Deleting connection...');

        try {
            await Excel.run(async (context) => {
                const settings = context.workbook.settings;
                const connectionsSetting = settings.getItemOrNullObject("connections");
                connectionsSetting.load("value");
                await context.sync();

                let connections = connectionsSetting.value ? JSON.parse(connectionsSetting.value) : [];
                connections = connections.filter(c => !(c.type === 'power-automate' && c.name === 'Send Personalized Email'));

                settings.add("connections", JSON.stringify(connections));
                await context.sync();

                setCurrentConnection(null);
                setUrl('');
                setStatus('Connection deleted successfully!');
                setTimeout(() => {
                    onClose();
                }, 1500);
            });
        } catch (error) {
            console.error('Error deleting connection:', error);
            setStatus('Error deleting connection');
        }
    };

    if (!isOpen) return null;

    return (
        <div className="fixed inset-0 bg-gray-600 bg-opacity-75 flex items-center justify-center z-50">
            <div className="bg-white rounded-lg shadow-xl p-6 w-full max-w-md">
                <div className="flex justify-between items-center mb-4">
                    <h3 className="text-lg font-semibold text-gray-800">Configure Power Automate</h3>
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

                {isLoading ? (
                    <div className="flex justify-center items-center py-8">
                        <div className="animate-spin rounded-full h-8 w-8 border-t-2 border-b-2 border-blue-600"></div>
                    </div>
                ) : (
                    <>
                        <div className="mb-4">
                            <label htmlFor="pa-url" className="block text-sm font-medium text-gray-700 mb-2">
                                Power Automate HTTP URL
                            </label>
                            <input
                                type="text"
                                id="pa-url"
                                value={url}
                                onChange={(e) => setUrl(e.target.value)}
                                className="w-full px-3 py-2 border border-gray-300 rounded-md shadow-sm focus:outline-none focus:ring-blue-500 focus:border-blue-500"
                                placeholder="https://prod-xx.eastus.logic.azure.com..."
                            />
                            <p className="mt-2 text-xs text-gray-500">
                                This URL is generated by the "When a HTTP request is received" trigger in your Power Automate flow.
                            </p>
                        </div>

                        <p className="text-xs text-center h-4 mb-4" style={{ color: status.includes('Error') ? '#dc2626' : '#16a34a' }}>
                            {status}
                        </p>

                        <div className="flex justify-between gap-2">
                            {currentConnection && (
                                <button
                                    onClick={handleDelete}
                                    className="px-4 py-2 bg-red-600 text-white rounded-md hover:bg-red-700"
                                >
                                    Delete
                                </button>
                            )}
                            <div className={`flex gap-2 ${currentConnection ? '' : 'ml-auto'}`}>
                                <button
                                    onClick={onClose}
                                    className="px-4 py-2 bg-gray-200 text-gray-700 rounded-md hover:bg-gray-300"
                                >
                                    Cancel
                                </button>
                                <button
                                    onClick={handleSave}
                                    className="px-4 py-2 bg-blue-600 text-white rounded-md hover:bg-blue-700"
                                >
                                    Save
                                </button>
                            </div>
                        </div>
                    </>
                )}
            </div>
        </div>
    );
}
