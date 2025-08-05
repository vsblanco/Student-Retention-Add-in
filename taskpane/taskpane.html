<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8" />
    <meta http-equiv="X-UA-Compatible" content="IE=Edge" />
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <title>Student Retention Add-in</title>

    <!-- Office JavaScript API -->
    <script type="text/javascript" src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"></script>

    <!-- ExcelJS for reading .xlsx files -->
    <script type="text/javascript" src="https://cdnjs.cloudflare.com/ajax/libs/exceljs/4.3.0/exceljs.min.js"></script>

    <!-- Tailwind CSS for styling -->
    <script src="https://cdn.tailwindcss.com"></script>

    <style>
        /* A little extra styling */
        body {
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
        }
    </style>
</head>
<body class="bg-gray-50">
    <!-- Main Content (Initially Hidden) -->
    <div id="main-content" class="p-4 hidden">
        <!-- Tab Navigation -->
        <div class="border-b border-gray-200">
            <nav class="-mb-px flex space-x-6" aria-label="Tabs">
                <button id="tab-details" type="button" class="whitespace-nowrap py-3 px-1 border-b-2 font-medium text-sm border-blue-500 text-blue-600" aria-current="page">
                    Student Details
                </button>
                <button id="tab-history" type="button" class="whitespace-nowrap py-3 px-1 border-b-2 font-medium text-sm border-transparent text-gray-500 hover:text-gray-700 hover:border-gray-300">
                    Student History
                </button>
            </nav>
        </div>

        <!-- Tab Content Panels -->
        <div class="pt-5">
            <!-- Student Details Panel -->
            <div id="panel-details">
                <div class="flex flex-col md:flex-row gap-4">
                    <!-- Left Column: Identity -->
                    <div class="w-full md:w-1/3 flex flex-col items-center text-center p-4 bg-white rounded-lg shadow-inner">
                        <div id="student-avatar" class="w-24 h-24 rounded-full flex items-center justify-center text-white font-bold text-4xl bg-gray-400 mb-3">
                            --
                        </div>
                        <h2 id="student-name-display" class="text-xl font-bold text-gray-800" title="Select a row">Select a row</h2>
                        <div class="mt-1">
                            <span id="assigned-to-badge" class="px-2 py-0.5 text-xs font-semibold text-gray-800 bg-gray-200 rounded-full">Unassigned</span>
                        </div>
                        <p class="mt-3 text-sm text-gray-600">ID: <span id="student-id-display">N/A</span></p>
                        <p class="mt-1 text-xs text-gray-500">Last LDA: <span id="last-lda-display">N/A</span></p>
                    </div>

                    <!-- Right Column: Info & Stats -->
                    <div class="w-full md:w-2/3 flex flex-col gap-4">
                        <!-- Stats Block -->
                        <div class="flex gap-4">
                            <div id="days-out-stat-block" class="flex-1 p-3 text-center rounded-lg bg-gray-200 text-gray-800">
                                <div id="days-out-display" class="text-3xl font-bold">--</div>
                                <div class="text-xs font-semibold">DAYS OUT</div>
                            </div>
                            <div id="grade-stat-block" class="flex-1 p-3 text-center rounded-lg bg-gray-200 text-gray-800 transition-colors duration-150">
                                <div id="grade-display" class="text-3xl font-bold">--%</div>
                                <div class="text-xs font-semibold">GRADE</div>
                            </div>
                        </div>
                        <!-- Contact Info -->
                        <div class="p-4 bg-white rounded-lg shadow-inner flex-grow">
                            <h3 class="font-bold text-gray-700 mb-3">Contact Information</h3>
                            <div class="space-y-1 text-sm">
                                <div id="copy-primary-phone" class="flex items-center p-2 rounded-md hover:bg-gray-100 cursor-pointer transition-colors duration-150">
                                    <svg class="w-5 h-5 text-gray-400 mr-3 flex-shrink-0" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M3 5a2 2 0 012-2h3.28a1 1 0 01.948.684l1.498 4.493a1 1 0 01-.502 1.21l-2.257 1.13a11.042 11.042 0 005.516 5.516l1.13-2.257a1 1 0 011.21-.502l4.493 1.498a1 1 0 01.684.949V19a2 2 0 01-2 2h-1C9.716 21 3 14.284 3 6V5z"></path></svg>
                                    <div class="flex-grow"><div class="text-xs text-gray-500">Primary Phone</div><div id="primary-phone-display" class="font-semibold">N/A</div></div>
                                    <span class="copy-feedback text-xs text-green-600 font-semibold hidden ml-2">Copied!</span>
                                </div>
                                <div id="copy-other-phone" class="flex items-center p-2 rounded-md hover:bg-gray-100 cursor-pointer transition-colors duration-150">
                                    <svg class="w-5 h-5 text-gray-400 mr-3 flex-shrink-0" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M3 5a2 2 0 012-2h3.28a1 1 0 01.948.684l1.498 4.493a1 1 0 01-.502 1.21l-2.257 1.13a11.042 11.042 0 005.516 5.516l1.13-2.257a1 1 0 011.21-.502l4.493 1.498a1 1 0 01.684.949V19a2 2 0 01-2 2h-1C9.716 21 3 14.284 3 6V5z"></path></svg>
                                    <div class="flex-grow"><div class="text-xs text-gray-500">Other Phone</div><div id="other-phone-display" class="font-semibold">N/A</div></div>
                                    <span class="copy-feedback text-xs text-green-600 font-semibold hidden ml-2">Copied!</span>
                                </div>
                                <div id="copy-student-email" class="flex items-center p-2 rounded-md hover:bg-gray-100 cursor-pointer transition-colors duration-150">
                                    <svg class="w-5 h-5 text-gray-400 mr-3 flex-shrink-0" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M3 8l7.89 5.26a2 2 0 002.22 0L21 8M5 19h14a2 2 0 002-2V7a2 2 0 00-2-2H5a2 2 0 00-2 2v10a2 2 0 002 2z"></path></svg>
                                    <div class="flex-grow"><div class="text-xs text-gray-500">Student Email</div><div id="student-email-display" class="font-semibold">N/A</div></div>
                                    <span class="copy-feedback text-xs text-green-600 font-semibold hidden ml-2">Copied!</span>
                                </div>
                                <div id="copy-personal-email" class="flex items-center p-2 rounded-md hover:bg-gray-100 cursor-pointer transition-colors duration-150">
                                    <svg class="w-5 h-5 text-gray-400 mr-3 flex-shrink-0" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M3 8l7.89 5.26a2 2 0 002.22 0L21 8M5 19h14a2 2 0 002-2V7a2 2 0 00-2-2H5a2 2 0 00-2 2v10a2 2 0 002 2z"></path></svg>
                                    <div class="flex-grow"><div class="text-xs text-gray-500">Personal Email</div><div id="personal-email-display" class="font-semibold">N/A</div></div>
                                    <span class="copy-feedback text-xs text-green-600 font-semibold hidden ml-2">Copied!</span>
                                </div>
                            </div>
                        </div>
                    </div>
                </div>
                <!-- Today's Outreach -->
                 <div class="mt-4 p-4 bg-white rounded-lg shadow-inner">
                    <h3 class="font-bold text-gray-700 mb-2">Today's Outreach</h3>
                    <p class="text-sm text-gray-500">No outreach logged today.</p>
                </div>
            </div>

            <!-- Student History Panel (Initially Hidden) -->
            <div id="panel-history" class="hidden space-y-4">
                <!-- New Comment Section -->
                <div class="p-4 bg-white rounded-lg shadow-inner">
                    <h3 class="font-bold text-gray-700 mb-2">Add New Comment</h3>
                    <textarea id="new-comment-input" class="w-full p-2 border border-gray-300 rounded-md" rows="3" placeholder="Type your comment here..."></textarea>
                    <button id="submit-comment-button" class="mt-2 w-full bg-blue-600 text-white font-bold py-2 px-4 rounded-lg hover:bg-blue-700 focus:outline-none focus:ring-2 focus:ring-blue-500 focus:ring-opacity-50 transition-colors duration-200 disabled:bg-gray-400" disabled>
                        Submit Comment
                    </button>
                    <p id="comment-status" class="text-xs text-gray-500 mt-2 h-4"></p>
                </div>
                <!-- History Display Section -->
                <div id="history-content" class="p-4 bg-white rounded-lg shadow-inner">
                    <p class="text-gray-500">Select a student row to see their history.</p>
                </div>
            </div>
        </div>
    </div>

    <!-- User Selection Modal (New) -->
    <div id="user-selection-modal" class="hidden fixed inset-0 bg-gray-600 bg-opacity-50 overflow-y-auto h-full w-full flex items-center justify-center z-50">
        <div class="relative p-5 border w-96 shadow-lg rounded-md bg-white">
            <div class="mt-3 text-center">
                <h3 class="text-lg leading-6 font-medium text-gray-900">Select User</h3>
                <div class="mt-2 px-7 py-3">
                    <p class="text-sm text-gray-500 mb-4">
                        Please select the user for this session.
                    </p>
                    <select id="user-selection-dropdown" class="w-full p-2 border border-gray-300 rounded-md">
                        <!-- Options will be populated by JS -->
                    </select>
                </div>
                <div class="items-center px-4 py-3">
                    <button id="confirm-user-button" class="px-4 py-2 bg-blue-500 text-white text-base font-medium rounded-md w-full shadow-sm hover:bg-blue-600 focus:outline-none focus:ring-2 focus:ring-blue-300">
                        Confirm
                    </button>
                </div>
            </div>
        </div>
    </div>
    
    <!-- Your Add-in's Code -->
    <script type="text/javascript" src="taskpane.js"></script>
    <script>
        // Debug log to confirm the taskpane HTML file is loaded.
        console.log("Taskpane HTML page loaded successfully.");
    </script>
</body>
</html>
