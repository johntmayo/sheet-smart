# Sheet Smart — Phase Progress

## Phase 1: Audit & Visibility (CURRENT)

**Status:** Complete and running

**Goal:** Get clear visibility into the current state of all ~55 spreadsheets vs the master.

**What's built:**
- `Code.gs` — standalone Apps Script that scans every sheet in a Drive folder and produces a four-tab audit report:
  - **Overview** — one row per spreadsheet with column counts, data rows, specific field counts (APN, Damage, For Sale, Sold Since Fire), missing/extra columns vs master, last edited timestamp, last editor
  - **Column Detail** — one row per column per spreadsheet showing position, whether it exists in master, and non-blank cell count
  - **Master Columns** — the canonical column list for reference
  - **Duplicate Resident IDs** — every `resident_ID` value that appears more than once across all sheets, with the spreadsheet name, link, row number, total occurrences, and whether the duplicates are within one sheet or across sheets
- Conditional formatting on the overview (light red for low counts, white-to-green gradient on data rows)
- Requires Drive Advanced Service enabled for last-editor lookup

**How to re-run:** Paste updated `Code.gs` into the Apps Script project and run `runAudit()`. Each run creates a fresh audit report.

---

## Phase 2: Lookup-and-Merge Tool

**Status:** In progress

**Goal:** A multi-operation data integration tool — import new data into the master, push data from the master out to the ~55 user sheets, and add missing columns to all user sheets.

**Interface:** Container-bound Apps Script inside a "Sheet Smart Config" Google Spreadsheet. Custom menu: `Sheet Smart` with six actions.

**Config spreadsheet layout:**
- **Settings tab:** Source Spreadsheet ID, Match Column (usually APN), Target Folder ID, Master Spreadsheet ID
- **Column Mapping tab:** Source Column | Target Column (one row per column to merge or add)

**Architecture (two .gs files, same Apps Script project):**
- `MergeEngine.gs` — core engine: `readMergeConfig_`, `buildSourceLookup_`, `mergeIntoTarget_`, `addColumnsToTarget_`, `appendToSyncLog_`, and supporting helpers
- `Corrections.gs` — UI layer: `onOpen` menu, sync entry points and dry-run wrappers, `setupConfigTabs`

**Menu actions (`Sheet Smart` menu):**
- **Import → Master** / **(Dry Run)** — External Source → Master; adds missing columns first, then fills blank cells; conflicts logged
- **Push → User Sheet** / **(Dry Run)** — Master → single user spreadsheet; adds missing columns first, then fills blank cells; conflicts logged
- **Push → User Sheets Folder** / **(Dry Run)** — Master → all sheets in folder; adds missing columns first, then fills blank cells; conflicts logged
- **Set Up Config Tabs** — formats Settings and Column Mapping tabs with labels, descriptions, and instructions; preserves existing values

**Settings keys (each operation reads only what it needs):**
- `Master Spreadsheet` — canonical master; destination for Import, source for Push
- `External Source` — outside data source for Import only (formerly `Source Spreadsheet`)
- `User Sheet` — single user sheet for Push → User Sheet only (new)
- `User Sheets Folder` — folder for Push → Folder only (formerly `Target Folder`)
- `Match Column` — shared by all operations

**Key behaviors:**
- Each Sync automatically adds any missing target columns before filling data — no separate step needed
- New columns are appended after the last existing column, formatted as plain General (no inherited checkboxes)
- Blank target cells are filled; non-blank differing cells are logged as conflicts and not overwritten
- Placeholder rows in Column Mapping (e.g. `(header name in master)`) and rows with either cell blank are silently skipped
- Unified sync log tab (`Last Sync - Master` or `Last Sync - Folder`) shows Column Added / Filled / Conflict / Error entries; cleared and rewritten on each run

**Open questions:**
- Should column order in targets match the master after columns are added?

---

## Phase 3: Ongoing Monitoring (Future)

**Ideas:**
- Scheduled trigger (daily/weekly) that re-runs the audit automatically
- Email summary when sheets drift out of sync
- `onEdit` triggers on each sheet to log changes to a central audit log (forward-looking edit history)
