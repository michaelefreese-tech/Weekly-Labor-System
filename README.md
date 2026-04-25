DMES Weekly Labor Tracker
A real-world labor tracking and validation system built for high-volume store operations using Google Sheets + Apps Script.

🚀 Live Demo
Open Demo Sheet
To test functionality: open the sheet and create your own copy.

📌 Project Status
Current Version: v1.0
Status: Stable / In-use (live store environment)

🔧 Core Capabilities

Tracks weekly labor across a 4-week period
Automates validation for all required inputs
Prevents invalid or incomplete data entry
Enforces structured workflows through sheet protection
Restores full sheet layout automatically if altered


⚙️ Key Features
🔒 Sheet Protection System

Full-sheet lock with controlled editable cells
Password-protected unlock system
Automatic restore on unauthorized edits

🧼 Restore Engine

Rebuilds full sheet structure:

Merges
Colors
Borders
Strings
Formulas


Prevents formatting corruption

✅ Validation System

Import validation (blocks bad data)
Live validation on COVER inputs
Visual feedback (red = incomplete, white = valid)

📅 4-Week Period System

Dynamic week generation
Automatic period tracking
Week navigation with persistent state

📥 Import Pipeline

Controlled import from raw data sheet
Week-based mapping into COVER
Prevents incomplete imports


🛠️ Built With

Google Apps Script
Google Sheets


💡 Why I Built This
Built from real operational experience managing high-volume stores.
Weekly labor tracking is prone to errors when handled manually. This system enforces structure, validates inputs, and protects the integrity of the data—reducing mistakes and improving consistency.

🚀 How to Use

Open the Google Sheet
Make a copy of the file
Run InitializeThisFile()
Set password when prompted
Use the IMPORT sheet to load data
Review results on the COVER sheet


📸 Screenshots

COVER Sheet (Labor Overview)

Displays weekly labor totals, variance, and structured reporting layout.

IMPORT Sheet (Data Input)

Used to input raw labor data before importing into the system.

Validation Feedback

Highlights missing or invalid inputs using real-time visual indicators.

Protected Sheet Behavior

Prevents unauthorized edits and restores structure automatically.

📜 Version History
v1.0 — 04.25.2026

Initial GitHub release
Full labor tracking system
Restore engine implementation
Validation system (import + live inputs)
Sheet protection framework
Week-based import pipeline


⚠️ Notes

This system depends on a specific spreadsheet structure (sheet names, ranges, and layout)
Code is provided as a reference implementation of system logic
Demo sheet is a sanitized version with no real store data


📫 Contact
📧 Email Me
Open to feedback, collaboration, and opportunities.
