# Specification: Dragon Boat Practice Workflow Automation

This document formalizes the requirements and implementation decisions for the Antigravity 2.0 Dragon Boat Practice Workflow Automation.

---

## Problem Statement

Paddling coaches design high-resolution training schedules (day's session) in Excel/Google Sheets workbooks containing standardized timings and intensities (gears). Currently, transforming these spreadsheet structures into mobile Tabata Timer JSON profiles and logging the training sessions chronologically in a master matrix is a tedious, manual, and error-prone transcription process. This leads to configuration errors, copy-paste leftovers (like template placeholder titles fanning across multiple files), and a lack of real-time visibility/parity between physical timer configurations and the actual master logs.

---

## Solution

The solution is an Antigravity 2.0 automated workflow pipeline. Coaches save their calculated spreadsheets in the workspace. The automation system:
1. Detects loose spreadsheets automatically or accepts a manual trigger path.
2. Parses the sheet deterministically using a local Python script.
3. Automatically infers objectives, tags, and summary strings.
4. Alerts the coach about any title discrepancies or copying remnants, allowing interactive overrides.
5. Standardizes and exports files to designated folders: spreadsheets go to `/01-Calculators_Sheets/`, and Tabata JSONs go to `/02-App_Timers_JSON/`.
6. Appends structured records to `sessions.json` which is displayed in an interactive, filterable client-side dashboard (`index.html`).
7. Supports a `/log-PP-only` workflow to log historical plans via text shorthand (omitting JSON generation and flagging files as missing in the UI).

---

## User Stories

1. As a coach, I want to drop a new spreadsheet into the workspace root, so that the pipeline automatically discovers and processes it.
2. As a coach, I want to run the pipeline with a specific file path, so that I can manually process files outside the root directory.
3. As a coach, I want the parser to dynamically handle cycle/set repetitions and expand them flatly in the JSON output, so that the mobile app plays the full sequences sequentially without looping bugs.
4. As a coach, I want the system to warn me if the spreadsheet contains conflicting title cells, so that I can fix template leftovers before they pollute the master log.
5. As a coach, I want to review and override the generated summary text and tags before they are committed, so that the log data is accurate.
6. As a coach, I want to use a command like `/log-PP-only` with text-shorthand, so that I can retroactively document historical workouts without generating unnecessary JSON config files.
7. As a paddler or coach, I want to view an interactive web page listing all sessions, so that I can sort them chronologically and filter them by objective tags (e.g. HIT + Race Ready).
8. As a coach, I want all my workbook and timer files to be cleanly named with the date and sanitized title, so that my archive remains perfectly organized.

---

## Implementation Decisions

### 1. No External AI Runtime Dependencies
* The pipeline script (`pipeline.py`) runs 100% locally and deterministically using standard Python libraries (`openpyxl` for spreadsheet reading, regex for title/date parsing). No paid API keys or local LLM runtimes are required for day-to-day timer generation.

### 2. Master Log Storage
* Metadata is stored in a clean `sessions.json` database.
* Example schema:
```json
[
  {
    "date": "Jul 06, 2026",
    "title": "The Over-Under",
    "summary": "7' G4 Warm-Up + 3' Rest + 7' Starts Drill + 3' Rest. Main Cycle: 6 sets of [3' @ G6 + 1' @ G7 + 1' @ G5 + 4.5' Rest]. Tabata: 8 sets of [20\" @ G10 + 10\" Rest]. 5' G0 Cooldown.",
    "tags": ["Endurance", "Race Ready"],
    "spreadsheet_path": "01-Calculators_Sheets/Jul 06, 2026 - The Over-Under.xlsx",
    "tabata_path": "02-App_Timers_JSON/Jul 06, 2026 - The Over-Under.tabata"
  }
]
```

### 3. Interactive UI
* A single static `index.html` file written in vanilla HTML, JS, and CSS reads the `sessions.json` at run-time, offering sorting, multi-select tag filtering, and dynamic HTML rendering.

### 4. Folder Structure
* Rigid spatial isolation:
  * `/01-Calculators_Sheets/` for spreadsheets.
  * `/02-App_Timers_JSON/` for Tabata JSON profiles.
  * Root directory contains the parser `pipeline.py`, `sessions.json`, and the dashboard `index.html`.

### 5. Date Extraction & File Naming
* Date is extracted from the spreadsheet filename using regex `([A-Za-z]{3} \d{1,2}, \d{4})` (matching MMM DD, YYYY).
* Filenames are renamed to `<Date> - <Sanitized Title>.<ext>` for consistency.

### 6. Tabata Generation Mapping
* Cell Type column values `"type": 0` (warmup/cooldown), `"type": 1` (work), `"type": 2` (rest).
* Minutes/Seconds values converted to integer seconds.
* Contiguous repeating blocks (cycles/sets) expanded to flat JSON objects, prefixing `"description"` with `Set X/Y - <Description>`.
* Randomize `workout.colorId` (integer 0-15) for variety.

### 7. Log-only Logic
* Ingests text strings (e.g., `3x(2'on 3'off,4'on 4'off)`) directly, saves to JSON with null/missing file indicators, updates the index, and skips JSON file generation.

---

## Testing Decisions

* **Good Test Principle**: Validate external behavior: file discovery, parsed time accuracy (seconds conversion), correct flat set multiplication, correct date-matching from filename, and JSON logging schema. Don't mock the filesystem unnecessarily; run real verification reads/writes on a mock directory.
* **Test Suite**: A local `test_pipeline.py` script verifying parsing logic against dummy spreadsheets.

---

## Out of Scope

* Integration with live Google Drive APIs or Google Sheets APIs for real-time cloud-sync. The tool operates purely on local files synced by Google Drive Desktop.
* Mobile app integration (loading JSON directly into the mobile app). The user must import the generated `.tabata` files into their Tabata Timer app manually.
