# Student Retention Excel Add-in

## Overview

This Excel Add-in is a powerful tool designed to help educational institutions improve student retention. It seamlessly integrates with Microsoft Excel to provide a comprehensive solution for tracking, analyzing, and engaging with students. The add-in is designed to save educators and administrators countless hours by automating repetitive tasks and providing actionable insights. It enhances team collaboration by allowing for shared notes and data.

## Features

* **Automatic Report Import**: Easily import student data from various report formats directly into Excel. The add-in is highly customizable, allowing you to tailor the import process and the display of your reports to your specific needs.
* **Student View**: A dedicated view that presents student data in a clear, concise, and easy-to-read format. This allows for a quick and holistic understanding of each student's status.
* **Collaborative Notes**: A built-in notes section allows you and your team to add, view, and edit notes for each student. All notes are securely saved within your Excel workbook, ensuring that everyone on the team has access to the latest information.
* **Submission Checker Chrome Extension (Optional)**: An optional Chrome extension that enhances the functionality of the add-in. It connects to the add-in via Pusher to:
    * Automatically open multiple tabs of students' gradebooks in Canvas.
    * Detect missing assignments and generate a report.
* **Personalized Emails**: A feature that allows you to send personalized emails to students based on the data in your workbook. This can be automated using Power Automate to send customized emails at scale, for example, to students with missing assignments or low grades.
* **Risk Index Calculation**: The add-in can calculate a "Risk Index" for each student based on customizable formulas. This helps in identifying at-risk students who may require additional support.
* **Analytics Dashboard**: Visualize student data with charts and graphs to identify trends and patterns.

## Getting Started

To get started with the Student Retention Add-in, you'll need to load it into your Excel application.

### Prerequisites

* Microsoft Excel 2016 or later

### Installation

1.  **Download the Add-in**: Download the add-in files from the GitHub repository.
2.  **Sideload the Add-in**: Follow the instructions provided by Microsoft to sideload the Office Add-in in your Excel application. You will need to use the `manifest.xml` file from this repository.

## Usage

Once the add-in is installed, you will see a new tab in your Excel ribbon. From there you can access all the features of the add-in.

1.  **Import Data**: Use the "Import Report" button to import your student data.
2.  **View Student Data**: Select a student to view their detailed information in the "Student View".
3.  **Add Notes**: Use the "Notes" section to add or view notes for the selected student.
4.  **Send Emails**: Use the "Send Personalized Email" feature to send customized emails to students.
5.  **Connect to Chrome Extension**: To use the Submission Checker functionality, install the optional Chrome extension and connect it to the add-in using the "Connections" tab and Pusher.

## Optional Chrome Extension

The optional Submission Checker Chrome extension adds powerful functionality to the add-in. It uses a robotic process to automate the checking of student submissions in Canvas.

### Extension Features

* **Automated Gradebook Checking**: Automatically opens multiple tabs of student gradebooks in Canvas to check for submissions.
* **Missing Assignment Detection**: Detects missing assignments for each student and generates a report.

### Installation

*The Chrome extension is not included in this repository and needs to be installed separately from the Chrome Web Store.* (Please provide a link to the Chrome extension here).

## Integrations

* **Pusher**: Used to connect the Excel Add-in with the Chrome extension for real-time communication.
* **Canvas LMS**: The Chrome extension integrates with Canvas to check for student submissions.
* **Microsoft Power Automate**: Can be used to automate the sending of personalized emails to students.

## Contributing

We welcome contributions to the Student Retention Add-in! If you have an idea for a new feature or have found a bug, please open an issue on our GitHub repository.
