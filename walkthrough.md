# Walkthrough - Dragon Boat Practice Workflow Automation

We have completed the implementation of the automated practice logging and Tabata JSON association lifecycle. 

Here is a summary of the accomplishments, components, and how to verify and run the system.

---

## 🛠️ Summary of Changes

### 1. Central Pipeline Script
* **[pipeline.py](file:///c:/Users/carlo/My%20Drive%20%28carlos.peralta.gutierrez@gmail.com%29/SNCC%20-%20Projects/DB-Practice%20Plan/Practices/pipeline.py)**: Built a local Excel workbook parser using `openpyxl`.
  * **Automatic Ingestion**: Scans the root directory for loose spreadsheets or accepts manual paths.
  * **Decoupled Tabata Association**: Decoupled automatic timer JSON generation. The script now automatically discovers any matching `.tabata` file (same name as the Excel file) or prompts you to link an existing file. If none is associated, it defaults to `null`.
  * **Intelligent Title & Date Parsing**: Merges Plan Header (`Practice Plan X:`) and Title Cell (`Practice:`) to resolve template leftovers. Extracts dates via regex from filenames.
  * **Heuristic Summary/Tags**: Generates standard sport summaries (e.g., `4 sets of [4' @ G5 + 3' Rest]`) and infers tags (`HIT`, `Race Ready`, etc.) dynamically from gears.
  * **Archiver & Cleanup**: Standardizes and relocates spreadsheets to `/01-Calculators_Sheets/` and associated Tabata JSONs to `/02-App_Timers_JSON/` using the format `<Date> - <Sanitized Title>`. Cleans up original loose files from the root.

### 2. Unified Master Log Database
* **[sessions.json](file:///c:/Users/carlo/My%20Drive%20%28carlos.peralta.gutierrez@gmail.com%29/SNCC%20-%20Projects/DB-Practice%20Plan/Practices/sessions.json)**: A local structured JSON file storing training metadata. Sorted chronologically with support for historical entries.

### 3. Interactive Web Dashboard
* **[index.html](file:///c:/Users/carlo/My%20Drive%20%28carlos.peralta.gutierrez@gmail.com%29/SNCC%20-%20Projects/DB-Practice%20Plan/Practices/index.html)**: A premium dark mode page with glassmorphic cards. Allows searching, multi-select tag filtering, date-sorting, and direct file access.
  * **Table Summary View**: Added a clean, Excel-like table view displaying columns for `Date`, `Title`, `Workout Summary`, `Tags`, and `Associated Files`, toggleable instantly via a top tab switcher.

---

## 🚀 How to Run and Verify

### A. Process a spreadsheet workbook (Calculators)
Save or copy any computed practice template spreadsheet (optionally alongside its matching `.tabata` JSON file with the same name) into the `Practices` workspace folder and run:
```powershell
python pipeline.py
```
* The script will scan the folder, extract the date, warn you if there are title inconsistencies (such as placeholder template names), and request confirmation for the summary and tags.
* It will search for a matching `.tabata` file. If not found, it will ask you if you wish to associate one.
* It moves the sheets to `/01-Calculators_Sheets/`, saves the Tabata profile to `/02-App_Timers_JSON/` (if linked), updates `sessions.json`, and cleans up loose files.

*To run in non-interactive mode (using defaults directly and logging N/A for missing tabata files):*
```powershell
python pipeline.py --yes
```

---

### B. Log historical sessions (Text Shorthand)
To retroactively catalog passed training sessions without creating Tabata files:
```powershell
python pipeline.py --log-only
```
It will prompt you to enter the Date, Workout Title, Shorthand Summary, and Objective Tags, updating the Master Log immediately.

---

### C. Open the Master Log Dashboard
Since web browsers restrict local JSON requests (`CORS`) on `file://` URLs, run a quick local web server in the `Practices` directory:
```powershell
python -m http.server 8000
```
Open [http://localhost:8000](http://localhost:8000) in your browser to view, search, and filter your practice history catalog in either Grid View or Table Summary.
