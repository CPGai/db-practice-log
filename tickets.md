# Tickets: Dragon Boat Practice Workflow Automation

A set of sequential, vertical tracer-bullet tickets to build the ingestion, parsing, reviewing, database, and dashboard systems for dragon boat training logs. Reference spec: [spec.md](file:///c:/Users/carlo/My%20Drive%20%28carlos.peralta.gutierrez@gmail.com%29/SNCC%20-%20Projects/DB-Practice%20Plan/Practices/spec.md)

Work the **frontier**: any ticket whose blockers are all done.

---

## 1. Core Parser Logic & Tabata JSON Export

**What to build:**
A local Python module that reads a practice template workbook, extracts timing rows, flatly expands repeating cycles/sets, and generates a valid Tabata Timer JSON profile.

**Blocked by:**
None — can start immediately.

* [ ] Implement an Excel reader in `pipeline.py` using `openpyxl` to parse type, time, unit, and description columns.
* [ ] Implement a flat set-repeater logic to duplicate interval rows (prepending `Set X/Y - ` to descriptions).
* [ ] Convert all time values (minutes/seconds/hours) to raw integer seconds.
* [ ] Write the structured result into the valid Tabata Timer JSON format with a randomized `workout.colorId` (0-15).
* [ ] Output the JSON onto a designated output file path.

---

## 2. Heuristic Summary & Tag Generation Engine

**What to build:**
A parser feature that analyzes parsed spreadsheet timings to generate a standard sport notation summary text and infer objective tags.

**Blocked by:**
1. Core Parser Logic & Tabata JSON Export

* [ ] Implement a summarizer function that collapses repeating consecutive timing sequences into a sport shorthand representation (e.g., `4x (3' @ G6 / 2'R)`).
* [ ] Implement a tag inference heuristic that scans gears (e.g. `G10` -> `HIT`) and comments (e.g. "Drill" -> `Technical paddling`) to assign tags.
* [ ] Ensure summary strings and tag lists match expected master log output structures.

---

## 3. Folder Ingestion & File Relocation

**What to build:**
An automated ingestion workflow that watches/scans the root directory for loose xlsx files, extracts dates, renames assets, and relocates them to their structured subdirectories.

**Blocked by:**
1. Core Parser Logic & Tabata JSON Export

* [ ] Implement a root folder scanner that detects loose `.xlsx` spreadsheets.
* [ ] Implement date extraction from filenames using regex (e.g., extracting `Jul 06, 2026` from `Jul 06, 2026.xlsx`).
* [ ] Implement file renaming to standard format: `<Date> - <Title>.<ext>`.
* [XY] Relocate source spreadsheets to `/01-Calculators_Sheets/` and Tabata JSONs to `/02-App_Timers_JSON/`.

---

## 4. Interactive Console Review Loop & Conflict Handling

**What to build:**
A terminal review prompts that warns the coach of title mismatches (especially placeholder defaults), displays inferred tags/summaries, and allows manual overrides.

**Blocked by:**
2. Heuristic Summary & Tag Generation Engine
3. Folder Ingestion & File Relocation

* [ ] Implement a checker that compares spreadsheet title cell values against sheet titles and filename headers, triggering a warning if they differ.
* [ ] Render the inferred tags and shorthand summary in the terminal console.
* [ ] Prompt the coach to accept, modify, or manually override the session title, tags, and summary before writing.
* [ ] Provide a mechanism to re-trigger updates if an incorrect title was previously committed.

---

## 5. JSON Log Database & `/log-PP-only` Mode

**What to build:**
A local database updater writing session entries to `sessions.json`, including a command-line fallback that logs text-shorthand workouts directly without generating Tabata files.

**Blocked by:**
4. Interactive Console Review Loop & Conflict Handling

* [ ] Maintain a unified `sessions.json` database containing dates, titles, summaries, tags, and filepaths.
* [ ] Implement `pipeline.py --log-only` workflow to ingest user text shorthand directly (e.g., `3x(2'on 3'off)`).
* [ ] Ensure the log-only mode records the workout metadata but leaves spreadsheet/tabata paths as null/missing.
* [ ] Save the updated log database safely back to disk.

---

## 6. Interactive Master Log Dashboard (`index.html`)

**What to build:**
A static, premium-designed HTML page in the root folder that reads `sessions.json` dynamically and displays all practices with tag filtering and sorting.

**Blocked by:**
5. JSON Log Database & `/log-PP-only` Mode

* [ ] Create `index.html` with vanilla JS, CSS, and HTML in a modern Dark Mode style.
* [ ] Read and render list entries dynamically from `sessions.json`.
* [ ] Add multi-select filtering for taxonomy tags (HIT, Endurance, etc.).
* [ ] Add ascending/descending sorting toggles for the session timeline.
* [ ] Display clickable direct download/access links for matching files, rendering a clean placeholder if files are marked as missing.
