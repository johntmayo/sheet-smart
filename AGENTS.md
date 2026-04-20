# Sheet Smart

Google Apps Script tooling for auditing and correcting ~55 Google Spreadsheets that were derived from a single master but have drifted over time. Owned and accessed via a Google service account.

## Project Structure

- `Code.gs` — standalone Apps Script (paste into script.google.com, not bound to a sheet)
- `MergeEngine.gs` — Phase 2 merge logic (container-bound; see Phase 2 architecture)
- `Corrections.gs` — Phase 2 custom menu / UI (container-bound)
- `PRD.md` — product requirements document (the "why" and "what")
- `phase_progress.md` — current roadmap and phase status
- `README.md` — setup instructions for end users

## How It Works

The script scans every spreadsheet in a Drive folder, reads the master spreadsheet's header row as the canonical column list, and compares each sheet against it. Output is a new Google Spreadsheet with three tabs: Overview, Column Detail, and Master Columns.

## Key Conventions

- **Apps Script only.** No Node/Python/build tools. Everything runs inside Google's script editor.
- **Standalone script.** Not bound to any spreadsheet. Do not use `SpreadsheetApp.getUi()` or other container-bound APIs.
- **Drive Advanced Service** is enabled in the project for `Drive.Files.get()` (last editor lookup). Reference it as `Drive`, not `DriveApp`, for advanced calls.
- **Batch reads/writes.** Always use `getDataRange().getValues()` to read and `setValues()` to write. Never loop `getValue()`/`setValue()` per cell.
- **Column counting patterns:**
  - `countNonBlankInColumn_` — counts cells that are not empty/null/undefined. Used for text/number columns like APN, Damage.
  - `countTrueInColumn_` — counts cells that are strictly `=== true`. Used for checkbox columns (Address - For Sale, Address - Sold Since Fire) because unchecked checkboxes return `false`, not blank.
  - `countUniqueAddressesMissingApn_` — counts unique Address values where the APN cell is blank.
- **Config at the top.** `FOLDER_ID` and `MASTER_SPREADSHEET_ID` are constants at the top of `Code.gs`. Never hardcode sheet IDs elsewhere.
- **Error handling.** If a sheet can't be opened, push an ERROR row to the overview and continue. Never let one bad sheet kill the whole run.

## Phase Plan

See `phase_progress.md` for full detail. Summary:
- **Phase 1 (current):** Audit — read-only visibility into column drift and data completeness
- **Phase 2 (next):** Bulk corrections — add missing columns, fill blank cells based on rules, log conflicts (non-blank cells with differing values) for manual review instead of overwriting
- **Phase 3 (future):** Ongoing monitoring — scheduled triggers, email alerts, onEdit audit logging

## Phase 2 architecture

Phase 2 uses two Apps Script files: **`MergeEngine.gs`** (lookup-and-merge core) and **`Corrections.gs`** (UI layer with a custom spreadsheet menu). Unlike Phase 1’s audit script, they are **container-bound** to a dedicated Google Spreadsheet named **“Sheet Smart Config”** (not a standalone script project).

The config spreadsheet holds:

- **Settings** — Source Spreadsheet, Match Column, Target Folder, Master Spreadsheet.
- **Column Mapping** — Source Column → Target Column pairs for aligning columns between source and targets.

**Three operations (each automatically adds missing columns then fills data):**

1. **Import → Master** — external source spreadsheet → master (by match column).
2. **Push → User Sheet** — master → a single user spreadsheet.
3. **Push → User Sheets Folder** — master → every spreadsheet in the configured folder.

A fourth menu item, **Set Up Config Tabs**, initializes the Settings and Column Mapping tabs with labels and instructions.

**Settings keys used by each operation:**
- Import → Master: `External Source`, `Master Spreadsheet`, `Match Column`
- Push → User Sheet: `Master Spreadsheet`, `User Sheet`, `Match Column`
- Push → User Sheets Folder: `Master Spreadsheet`, `User Sheets Folder`, `Match Column`

Each operation reads only its own settings; unused settings are ignored.

**Write policy (Option D):** fill blank cells only; when a cell already has a value that differs from the incoming value, **do not overwrite** — record it as a conflict for manual review. Never silently overwrite non-blank disagreements.

**Outputs:** run summaries, metrics, and unified sync logs (Column Added / Filled / Conflict / Error) are written to **tabs on the same config spreadsheet** (`Last Import`, `Last Push - User Sheet`, `Last Push - Folder`, and dry-run variants).

**UI:** `SpreadsheetApp.getUi()` and other container-bound APIs **are** used in `Corrections.gs` because the script is bound to the config spreadsheet. That restriction applies only to the **standalone** Phase 1 audit (`Code.gs`).

## When Editing Code.gs

- The overview header array, the data row arrays, and the error/empty row arrays must all have the same number of elements. When adding a column, update all three places plus the header in `writeReport_`.
- Conditional formatting column references (G, H, I, J, etc.) in `applyConditionalFormatting_` must be updated if columns are inserted before them.
- Helper functions use trailing underscore naming (`readMasterHeaders_`, `collectSpreadsheets_`, etc.) per Apps Script private function convention.
