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
Prevents incomplete or incorrect data entry
Enforces structured workflows through protection
Restores full sheet layout automatically


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


🛠 Built With

Google Apps Script
Google Sheets


💡 Why I Built This
Managing labor weekly requires accuracy and consistency.
Manual systems break easily.
This tool enforces structure, validates inputs, and protects data integrity—reducing mistakes and improving efficiency in real store operations.

🚀 How to Use

Make a copy of the sheet
Run InitializeThisFile()
Set password
Import weekly data
Review results on COVER


📸 Screenshots
(Add screenshots here)

📜 Version History
v1.0 — 04.25.2026

Initial GitHub release
Full labor tracking system
Restore engine
Validation system
Sheet protection framework
Import pipeline


⚠️ Notes

Depends on a fixed sheet structure
Demo is sanitized (no real data)
Code reflects production usage logic


📫 Contact
📧 michael.e.freese.tech@gmail.com
Open to feedback, collaboration, and opportunities.
