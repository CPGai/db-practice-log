# Implementation Plan - Dragon Boat Practice Workflow Automation

This plan outlines the architecture and execution steps for implementing the automated lifecycle pipeline within the `DB-Practice Plan` environment. 

## Goals
* **Automated Ingestion**: Automatically detect loose `.xlsx` spreadsheets in the workspace root or accept explicit file paths.
* **Intelligent Parsing**: Convert spreadsheet rows into structured cycles/sets and generate a production-ready `.tabata` JSON timer profile.
* **Review/Warning Loop**: Check for title inconsistencies, infer tags and summaries, and present a review step to the coach before final execution.
* **File Archival & Sorting**: Relocate spreadsheets to `/01-Calculators_Sheets/` and output JSONs to `/02-App_Timers_JSON/` using a standardized `<Date> - <Title>` naming format.
* **Historical Ingestion (`/log-PP-only`)**: Allow logging of text-only shorthand workout summaries to the Master Log without generating JSON timer files.
* **Master Log Dashboard**: Maintain a structured `sessions.json` database and render an interactive, high-density, filterable `index.html` dashboard.

---

## User Review Required

We have aligned on the following core design decisions during our `/grilling` session:
1. **Master Log Architecture**: Using a structured `sessions.json` file as the data store and a static `index.html` dashboard that renders and filters the data dynamically in the browser.
2. **Ingestion Trigger**: Automatic directory scanning for loose `.xlsx` files in the root folder, with a manual filepath override command, plus a `/log-PP-only` mode for text-based historical logs.
3. **Title Selection**: Option C (triggering an interactive review prompt if the title cell conflicts with the section header, allowing manual correction).
4. **Tag & Summary Inference**: Automatic background generation of objective tags and summary text, presented to you in the review step for confirmation or manual override.
5. **Standardized Naming**: Naming files as `<Date> - <Sanitized Title>.<extension>` (extracting the date from the filename via regex).

> [!IMPORTANT]
> **Pending User Action Item:**
> You need to set up a clean, standardized Excel/Sheets template and establish a workflow discipline of avoiding file duplication or partial edits. This prevents dirty/unrelated rows from confounding the parser.

---

## Open Questions

> [!NOTE]
> None currently. All major architectural directions were resolved during the grilling session.

---

## Proposed Changes

### Automation Core

#### [NEW] [pipeline.py](file:///c:/Users/carlo/My%20Drive%20%28carlos.peralta.gutierrez@gmail.com%29/SNCC%20-%20Projects/DB-Practice%20Plan/Practices/pipeline.py)
A central Python orchestration script that manages the entire ingestion, parsing, user review, file sorting, and master log update pipeline. Key features:
- **Scan & Ingest**: Detects loose xlsx files in the root, or accepts a command-line filepath / text-shorthand input.
- **XLSX Parser**: Reads the `Plan` sheet, groups cycles based on `sets`/`min`/`sec` indicators, and expands loops/sets dynamically.
- **Heuristic Engine**: Generates the structural workout summary and infers tags (e.g., *Endurance*, *HIT*) based on interval duration and gears.
- **Review UI**: Prompts the user to confirm/override the title, tags, and summary.
- **Tabata Generator**: Incorporates the Tabata-Timer JSON rules to produce the output configuration object.
- **File Organizer**: Creates target folders, standardizes filenames to `<Date> - <Title>`, and moves files.
- **JSON Logger**: Appends/updates entries in `sessions.json`.

---

### Master Log Database & UI

#### [NEW] [sessions.json](file:///c:/Users/carlo/My%20Drive%20%28carlos.peralta.gutierrez@gmail.com%29/SNCC%20-%20Projects/DB-Practice%20Plan/Practices/sessions.json)
The structured JSON database storing all training session metadata. Example schema:
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

#### [NEW] [index.html](file:///c:/Users/carlo/My%20Drive%20%28carlos.peralta.gutierrez@gmail.com%29/SNCC%20-%20Projects/DB-Practice%20Plan/Practices/index.html)
An interactive, high-density HTML dashboard using Vanilla HTML, CSS (Curated Dark Mode), and JS:
- Dynamically loads `sessions.json`.
- Implements tag-based filters (multi-select) and ascending/descending chronological sorting.
- Displays clickable direct links to the spreadsheet calculators and `.tabata` files.
- Gracefully flags missing assets (e.g., "Not saved/Historical-only" for text-only logs).

---

## Verification Plan

### Automated Tests
We will write a verification test suite in `scratch/test_pipeline.py` that verifies:
- Parser correctly maps minutes and seconds to Tabata JSON times.
- Loop/set expansion correctly prepends `Set x/y` and duplicates intervals.
- The regex correctly extracts dates from filenames (e.g., `Jul 06, 2026.xlsx`, `Exit Focus Mar 02, 2026.xlsx`).
- `/log-PP-only` successfully updates `sessions.json` without creating files.

### Manual Verification
1. Run `python pipeline.py` on an existing spreadsheet (e.g., copy `Jul 06, 2026.xlsx` to the root folder).
2. Verify the warning prompt triggers for title confirmation.
3. Check that the files are properly placed in `/01-Calculators_Sheets/` and `/02-App_Timers_JSON/`.
4. Open the generated `index.html` in a web browser to verify styling, filtering, and tag selection.
