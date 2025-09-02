<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8" />
    <meta http-equiv="X-UA-Compatible" content="IE=Edge" />
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <title>Connections</title>

    <!-- Office JavaScript API -->
    <script type="text/javascript" src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"></script>
    
    <!-- Pusher JavaScript Library (Local) -->
    <script src="pusher.min.js"></script>

    <!-- Tailwind CSS -->
    <script src="https://cdn.tailwindcss.com"></script>

    <!-- Custom CSS -->
    <style>
        .glass-modal-bg {
            backdrop-filter: blur(8px);
            -webkit-backdrop-filter: blur(8px);
            background-color: rgba(255, 255, 255, 0.6);
        }
        #log-header { cursor: pointer; }
        #log-arrow { transition: transform 0.2s; }
        #log-container.hidden { display: none; }
        .status-dot {
            width: 10px;
            height: 10px;
            border-radius: 50%;
            display: inline-block;
            margin-right: 8px;
            flex-shrink: 0;
        }
        .status-dot.disconnected { background-color: #9ca3af; } /* gray-400 */
        .status-dot.connecting { background-color: #f59e0b; } /* amber-500 */
        .status-dot.connected { background-color: #22c55e; } /* green-500 */
        .status-dot.error { background-color: #ef4444; } /* red-500 */
    </style>
</head>
<body class="bg-gray-100 font-sans">
    <div class="p-6 h-screen flex flex-col">
        <!-- Header -->
        <div class="flex justify-between items-start">
            <div>
                <h1 class="text-2xl font-bold text-gray-800">Connections</h1>
                <p class="mt-2 text-sm text-gray-600">Connect this add-in with other applications/services.</p>
            </div>
            <button id="new-connection-button" class="bg-blue-600 text-white font-bold py-2 px-4 rounded-lg hover:bg-blue-700 transition-colors duration-200 shadow-md flex items-center justify-center gap-2 whitespace-nowrap">
                <svg xmlns="http://www.w3.org/2000/svg" class="h-5 w-5" viewBox="0 0 20 20" fill="currentColor">
                  <path fill-rule="evenodd" d="M10 5a1 1 0 011 1v3h3a1 1 0 110 2h-3v3a1 1 0 11-2 0v-3H6a1 1 0 110-2h3V6a1 1 0 011-1z" clip-rule="evenodd" />
                </svg>
                New Connection
            </button>
        </div>

        <!-- Active Connections List -->
        <div class="mt-8 pt-6 border-t border-gray-200 flex-grow overflow-y-auto">
            <h2 class="text-lg font-bold text-gray-700 mb-4">Active Connections</h2>
            <div id="connections-list-container" class="space-y-4">
                <p id="no-connections-message" class="text-sm text-gray-500 italic">No connections have been added yet.</p>
            </div>
        </div>

        <!-- Debug Log Section -->
        <div id="debug-log-section" class="mt-6 pt-6 border-t border-gray-200">
             <div id="log-header" class="flex justify-between items-center cursor-pointer">
                <h2 class="text-lg font-bold text-gray-700">Debug Log</h2>
                <div class="flex items-center">
                    <button id="clear-log-button" class="text-xs text-blue-600 hover:underline mr-2">Clear Log</button>
                    <svg id="log-arrow" class="w-5 h-5 text-gray-600" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M19 9l-7 7-7-7"></path></svg>
                </div>
            </div>
            <div id="log-container" class="hidden mt-2 p-3 h-48 bg-gray-800 text-white font-mono text-xs rounded-lg overflow-y-auto border border-gray-600">
            </div>
        </div>
    </div>

    <!-- Modals -->
    <!-- Select Service Modal -->
    <div id="select-service-modal" class="hidden fixed inset-0 bg-black bg-opacity-40 flex items-center justify-center p-4 z-40">
        <div class="glass-modal-bg p-8 rounded-2xl shadow-2xl w-full max-w-lg border border-white/20">
            <h2 class="text-xl font-bold text-gray-800 mb-2 text-center">Add a New Connection</h2>
            <p class="text-sm text-gray-600 mb-6 text-center">Select a service to connect to.</p>
            <div class="grid grid-cols-1 md:grid-cols-2 gap-6">
                <!-- Power Automate Card -->
                <div class="bg-white p-6 rounded-lg shadow-md border border-gray-200 opacity-50 cursor-not-allowed">
                    <h3 class="text-lg font-semibold text-gray-800">Power Automate</h3>
                    <p class="text-sm text-gray-500 mt-2">Automate workflows between your favorite apps and services.</p>
                    <button class="w-full mt-4 bg-gray-400 text-white font-bold py-2 px-4 rounded-lg cursor-not-allowed">Coming Soon</button>
                </div>
                <!-- Pusher Card -->
                <button id="select-pusher-button" class="bg-white p-6 rounded-lg shadow-md border border-gray-200 hover:border-blue-500 hover:shadow-lg transition-all text-left">
                    <h3 class="text-lg font-semibold text-gray-800">Pusher</h3>
                    <p class="text-sm text-gray-500 mt-2">Enable real-time features like live submission highlighting.</p>
                    <div class="w-full mt-4 bg-blue-600 text-white font-bold py-2 px-4 rounded-lg text-center">
                        Configure
                    </div>
                </button>
            </div>
            <div class="mt-6 flex justify-end">
                <button id="cancel-select-service-button" class="px-4 py-2 rounded-lg bg-gray-200 text-gray-800 font-semibold hover:bg-gray-300 transition-colors">Cancel</button>
            </div>
        </div>
    </div>

    <!-- Pusher Configuration Modal -->
    <div id="pusher-config-modal" class="hidden fixed inset-0 bg-black bg-opacity-40 flex items-center justify-center p-4 z-50">
        <div class="glass-modal-bg p-8 rounded-2xl shadow-2xl w-full max-w-md border border-white/20">
            <h2 class="text-xl font-bold text-gray-800 mb-4">Configure Pusher Connection</h2>
            <div class="space-y-3 bg-white p-4 rounded-lg shadow-sm border border-gray-200">
                <div>
                    <label for="connection-name" class="block text-sm font-medium text-gray-700">Connection Name</label>
                    <input type="text" id="connection-name" class="mt-1 block w-full px-3 py-2 border border-gray-300 rounded-md" placeholder="e.g., Submission Checker">
                </div>
                <div>
                    <label for="pusher-key" class="block text-sm font-medium text-gray-700">App Key</label>
                    <input type="text" id="pusher-key" class="mt-1 block w-full px-3 py-2 border border-gray-300 rounded-md" placeholder="Your Pusher App Key">
                </div>
                <div>
                    <label for="pusher-secret" class="block text-sm font-medium text-gray-700">App Secret</label>
                    <input type="password" id="pusher-secret" class="mt-1 block w-full px-3 py-2 border border-gray-300 rounded-md" placeholder="Your Pusher App Secret">
                </div>
                <div>
                    <label for="pusher-cluster" class="block text-sm font-medium text-gray-700">Cluster</label>
                    <input type="text" id="pusher-cluster" class="mt-1 block w-full px-3 py-2 border border-gray-300 rounded-md" placeholder="e.g., us2, eu">
                </div>
                 <div class="mt-2 p-3 bg-yellow-50 border border-yellow-200 rounded-md">
                    <p class="text-xs text-yellow-800"><strong>Security Note:</strong> The App Secret is stored within your Excel workbook.</p>
                </div>
            </div>
            <div class="mt-6 flex justify-end gap-2">
                <button id="cancel-pusher-config-button" class="px-4 py-2 rounded-lg bg-gray-200 text-gray-800 font-semibold hover:bg-gray-300">Cancel</button>
                <button id="create-pusher-connection-button" class="px-4 py-2 rounded-lg bg-green-600 text-white font-bold hover:bg-green-700">Create</button>
            </div>
        </div>
    </div>

     <!-- Add Action Modal -->
    <div id="add-action-modal" class="hidden fixed inset-0 bg-black bg-opacity-40 flex items-center justify-center p-4 z-50">
        <div class="glass-modal-bg p-8 rounded-2xl shadow-2xl w-full max-w-lg border border-white/20">
            <h2 class="text-xl font-bold text-gray-800 mb-2 text-center">Add Action</h2>
            <p class="text-sm text-gray-600 mb-6 text-center">Select and configure an action for this connection.</p>
            
            <!-- Live Submission Highlighting Action -->
            <div class="bg-white p-4 rounded-lg shadow-md border border-gray-200">
                <h3 class="text-base font-semibold text-gray-800">Live Submission Highlighting</h3>
                <p class="text-sm text-gray-500 mt-1 mb-4">Highlights a student's row in real-time when a new submission is detected.</p>
                <div class="space-y-3">
                    <div>
                        <label for="pusher-channel" class="block text-sm font-medium text-gray-700">Channel Name</label>
                        <input type="text" id="pusher-channel" class="mt-1 block w-full px-3 py-2 border border-gray-300 rounded-md" placeholder="e.g., private-my-channel">
                    </div>
                    <div>
                        <label for="pusher-event" class="block text-sm font-medium text-gray-700">Event Name</label>
                        <input type="text" id="pusher-event" class="mt-1 block w-full px-3 py-2 border border-gray-300 rounded-md" placeholder="e.g., new-submission-found">
                    </div>
                </div>
                 <button data-action-type="liveHighlight" class="w-full mt-4 bg-blue-600 text-white font-bold py-2 px-4 rounded-lg hover:bg-blue-700">Add This Action</button>
            </div>

            <div class="mt-6 flex justify-end">
                <button id="cancel-add-action-button" class="px-4 py-2 rounded-lg bg-gray-200 text-gray-800 font-semibold hover:bg-gray-300 transition-colors">Cancel</button>
            </div>
        </div>
    </div>

    <script type="text/javascript" src="connections.js"></script>
</body>
</html>

