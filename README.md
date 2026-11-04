# Student Work Anonymizer

Google Apps Script for a teacher's Sheet: removes student names from Canvas submissions before an AI helps grade them, then sends grades back to Canvas.

This repo holds **code only**. No student names, work, IDs, grades or passwords are ever stored here. Student data stays in the teacher's school Google account and Canvas.

- `Starter.js`: pasted once into Code.gs in Apps Script. It builds the menu, loads the tool's code from a file in the teacher's school Drive, and installs updates from here (download only).
- `Code.js`: the tool itself. Install downloads it from here into that Drive file.
- `version.json`: the newest version number and a one-line note shown in the app.
- `appsscript.json`: the manifest (permissions) for a fresh install.
