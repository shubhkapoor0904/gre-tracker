# TargetMS — GRE / MS University Application Deadline Tracker

TargetMS is a complete, polished, and **100% FREE** personal application-management dashboard for tracking graduate school (MS/PhD) applications, GRE requirements, document checklists, deadlines, and calendar reminders.

---

## 🌟 Key Features

1. **Intelligent Excel Import (`Colleges.xlsx`)**:
   - Automatically attempts to load and parse `Colleges.xlsx` on first load using SheetJS.
   - Includes a drag-and-drop file uploader to import any updated `.xlsx`, `.xls`, or `.csv` file.
   - Dynamic header detection automatically maps university, program, degree, country, deadline, status, fee, and notes columns.
   - Preserves multiple programs for the same university as separate, independent application records.

2. **Executive SaaS Dashboard & Live Countdown**:
   - **Metrics Bar**: Instant stats on Total Applications, Submitted, In Progress, Unstarted, Critical/Overdue (<14 days), and High Priority.
   - **Timeline Carousel**: Chronologically sorted upcoming deadlines with real-time countdown pills (`82d SAFE`, `12d URGENT`, `3d CRITICAL`, `DUE TODAY`, `OVERDUE`).
   - **Modern SaaS Aesthetics**: Deep dark glassmorphism, Inter/Outfit typography, clean visual hierarchy, and gold priority accents.

3. **Dynamic Requirements Checklist & Smart Math**:
   - Tracks 11 core application requirements: Resume/CV, SOP, LOR 1-3, Transcripts, GRE, TOEFL/IELTS, App Form, App Fee, Final Submission.
   - **Smart Math Logic**: If a program waives or does not require GRE, the GRE requirement is automatically excluded from the completion % denominator so it doesn't penalize your progress.
   - Custom document items can be added per program (e.g. Portfolio, Video Essay, Diversity Statement).

4. **Smart Warnings Engine**:
   - Automated alerts banner for:
     - Deadline in $\le 14$ days with incomplete SOP or LORs.
     - Deadline in $\le 14$ days with status "Not Started".
     - Application marked "Ready to Apply" with $< 80\%$ document completion.
     - Passed / overdue deadlines.

5. **100% Free Calendar Sync & Reminders**:
   - **1-Click Google Calendar Event Links**: Instantly create deadline events with title, details, and portal URLs.
   - **Downloadable `.ics` iCalendar Files**: Single application or bulk all-deadlines `.ics` export with built-in alarm triggers (30d, 14d, 7d, 3d, 1d before).
   - **Browser Desktop Notifications**: Native Web Notification API alerts for pending deadlines.
   - **Free Google Apps Script Automation**: Provision of a copy-paste Apps Script to send free daily email reminders directly from Gmail.

6. **Search, Filter, Sort & Full Management (CRUD)**:
   - Full-text live search across universities, programs, notes, and countries.
   - Multi-field filters: Status, Urgency level, Priority (High/Med/Low), GRE requirement.
   - Multi-column sorting: Deadline (Earliest/Latest), Progress %, Priority, University Name (A-Z).
   - Add new university/program manually without touching code.
   - Detailed slide-over drawer for quick editing, notes, and document toggles.

7. **Export & Data Persistence**:
   - All data persists automatically in browser `localStorage`.
   - Export active tracker data to Excel (`.xlsx`) or `.csv`.
   - Full JSON backup export and restore capabilities.

---

## 🚀 How to Run

1. Simply open [`index.html`](file:///d:/gre%20tracker/index.html) in any modern web browser (Google Chrome, Microsoft Edge, Firefox, Safari).
2. If served via a local web server (e.g. `npx serve` or Live Server extension), `index.html` will automatically read `Colleges.xlsx` in the workspace directory.
3. You can also click **"Import Excel / CSV"** at the top right of the application header to upload `Colleges.xlsx` at any time!

---

## 📊 Summary of Initial Pre-Verified Programs

| University | Program | Deadline | Status | GRE Requirement |
| :--- | :--- | :--- | :--- | :--- |
| **Purdue University** | MS CS | Dec 1, 2026 | Documents Pending | Optional |
| **UIUC** | MS CS | Dec 15, 2026 | Documents Pending | Not Required |
| **Carnegie Mellon (CMU)** | MS CS | Dec 10, 2026 | Researching | Optional |
| **Georgia Tech** | MS CS | Feb 1, 2027 | Documents Pending | Optional |
| **Columbia University** | MS CS | Jan 15, 2027 | Not Started | Optional |
| **NYU (CDS)** | MS DS | Jan 22, 2027 | Documents Pending | Optional |
| **USC** | MS CS | Dec 15, 2026 | Not Started | Waived |
| **Northeastern** | MS CS | Apr 15, 2027 | Ready to Apply | Not Required |

---

## 📅 One-Time Setup for Email Reminders (Google Apps Script)

If you would like automated daily email notifications sent directly to your Gmail inbox:
1. Open [Google Drive](https://drive.google.com).
2. Click **New** > **More** > **Google Apps Script**.
3. Click the **"Calendar Sync & Alerts"** button in TargetMS header and copy the free script snippet.
4. Paste into the Apps Script editor and click **Save**.
5. Set a daily trigger under **Triggers** (Clock icon) $\rightarrow$ **Add Trigger** $\rightarrow$ **Time-driven** (Day timer: 8am to 9am).
6. That's it! 100% free daily email reminders will be sent to your Gmail whenever deadlines approach!
