# Sheet Smart — How to Use

Google Apps Script tooling for auditing and correcting ~55 Google Spreadsheets derived from a single master. Two independent tools: a **Phase 1 audit script** (`Code.gs`) and a **Phase 2 sync tool** (`MergeEngine.gs` + `Corrections.gs` + `Sidebar.html` in a config spreadsheet).

---

## Phase 1: Audit Report

Scans every spreadsheet in a Drive folder, compares columns against the master, and produces a structured report.

### What the report contains

| Tab | One row per… | Shows you |
|---|---|---|
| **Overview** | spreadsheet | column count, row count, zone, missing/extra row counts, missing/extra columns vs master, last editor, last edited |
| **Column Detail** | column × spreadsheet | column name, position, whether it exists in master, non-blank cell count |
| **Master Columns** | column in master | the canonical column list for reference |
| **Duplicate Resident IDs** | duplicate resident_id occurrence | which sheets contain the same resident_id, and whether duplication is within or across sheets |
| **Missing Rows** | resident expected in a sheet but absent | which captain's sheet is missing a resident master assigns to their zone |
| **Extra Rows** | resident present in a sheet but not expected | which captain's sheet has a resident that doesn't belong, with reason (Not in master / Wrong zone) |

### Setup (one-time)

1. Go to [script.google.com](https://script.google.com) and create a new **standalone** project (not bound to any spreadsheet).
2. Paste the contents of `Code.gs` into the editor.
3. Edit the two constants at the top of the file:
   - `FOLDER_ID` — the ID at the end of your Drive folder URL (`https://drive.google.com/drive/folders/`**`ABC123`**).
   - `MASTER_SPREADSHEET_ID` — the ID in your master sheet URL (`https://docs.google.com/spreadsheets/d/`**`XYZ789`**`/edit`).
4. Enable the **Drive Advanced Service**: click **Services** (+ icon in the left sidebar) → find **Drive API** → click **Add**.
5. Click **Run** → select `runAudit` → **Run**. Approve the permission prompt on first run.

### What the audit expects in the master and user sheets

For the Missing Rows / Extra Rows tabs to be meaningful, the master spreadsheet's first tab must contain:

- A `resident_id` column (unique identifier per row).
- A `ZoneName` column (which zone each resident belongs to).

Each user sheet should also contain a `ZoneName` column (normally pushed from master). The audit infers each user sheet's assigned zone by taking the most common non-blank value in its `ZoneName` column. Sheets without a `ZoneName` column (or with all blanks) are shown as `(no zone detected)` and excluded from row-membership diffing.

The `Resident Name`, `House`, and `Street` columns are used for display context in the report — useful but not required.

### Running it

Run `runAudit()` any time you want a fresh snapshot. Each run creates a new report spreadsheet and logs its URL in the execution log (`View → Logs`).

---

## Phase 2: Sync Tool

Operations for managing data and schema across your master spreadsheet and user sheets. Sync operations automatically add any missing columns before filling data — you never need a separate step.

### Core concepts

| Term | Means |
|---|---|
| **Master** | Your canonical spreadsheet. The source of truth for column structure and shared data. |
| **User Sheet** | One of the ~55 individual spreadsheets delivered to users. |
| **External Source** | Any other spreadsheet you want to import data FROM into the master (e.g. a sales feed, a contractor's file). |
| **Source Tab Name** | Optional tab inside the external source spreadsheet. If blank, imports use the first tab. |
| **Match Column** | The column that exists in both master and user sheets to match rows — usually APN. |
| **Workflow Preset** | A named, task-oriented configuration used by the guided sidebar (for example, "Update Master From Sales Tracker"). Presets can include per-column import policies. |

### The operations

| Menu item | FROM | TO | When to use |
|---|---|---|---|
| **Import → Master** | External Source | Master | New data has arrived in an outside sheet and you want it in the master |
| **Push → User Sheet** | Master | One user sheet | Push master data to a specific user sheet (fills blank cells in existing rows) |
| **Push → User Sheets Folder** | Master | All sheets in a folder | Push master data to all user sheets (fills blank cells in existing rows) |
| **Push Missing Rows → User Sheet** | Master | One user sheet | Append rows from master whose `ZoneName` matches the sheet's zone but whose `resident_id` isn't there yet |
| **Push Missing Rows → User Sheets Folder** | Master | All sheets in a folder | Same as above, but across every sheet in the folder |
| **Pull Missing Rows ← User Sheet** | One user sheet | Master | Append captain-created rows whose `resident_id` is not already in master |
| **Pull Missing Rows ← User Sheets Folder** | All sheets in a folder | Master | Same as above, but across every sheet in the folder |
| **Pull Data ← User Sheet** | One user sheet | Master | Pull captain-entered values into existing master rows using `Pull Column Policy`; also appends missing rows |
| **Pull Data ← User Sheets Folder** | All sheets in a folder | Master | Same as above, but across every sheet in the folder |
| **Rename Columns → User Sheets Folder** | Settings (`Rename Column - From`, `Rename Column - To`) | All sheets in a folder | Rename one header across all user sheets without touching row data |

> The regular **Push** operations only fill blank cells inside existing rows — they never add rows. The **Push Missing Rows** and **Pull Missing Rows** operations only add rows — they never modify existing ones. Use Push and Pull together to keep user sheets and master aligned.

> **Pull Data** is the broader captain-data import. It can fill blanks, overwrite existing master values, or log conflicts based on the `Pull Column Policy` tab.

### Setup (one-time)

1. Create a new Google Spreadsheet named **Sheet Smart Config**.
2. Open **Extensions → Apps Script**. Paste `MergeEngine.gs`, `Corrections.gs`, and `Sidebar.html` from this repo into separate script files in the same project. Save.
3. Reload the Sheet Smart Config spreadsheet. The **Sheet Smart** menu will appear.
4. Click **Sheet Smart → Set Up Config Tabs** to format the Settings, Column Mapping, Pull Column Policy, and Workflow Presets tabs.
5. Fill in the Settings tab (column B):

   | Setting | What to put here |
   |---|---|
   | Master Spreadsheet | ID of your master spreadsheet |
   | External Source | ID of an outside data source (only needed for Import → Master) |
   | Source Tab Name | Optional tab name inside the external source spreadsheet (e.g. `Sales Rollup by APN`) |
   | User Sheet | ID of a single user sheet to push to or pull from |
   | User Sheets Folder | ID of the Drive folder of user sheets |
   | Match Column | Column header used to match rows — e.g. `resident_id` |
   | Rename Column - From | Existing header name to rename (only needed for Rename Columns → User Sheets Folder) |
   | Rename Column - To | New header name to write (only needed for Rename Columns → User Sheets Folder) |
   | Sensitive Columns | Comma-separated list of column headers considered privacy-sensitive (e.g. `Person Notes, Contact Notes, Address Notes`). Only used by the Push Missing Rows operations — rows whose master record has a value in any of these columns are flagged for review. |

   You only need to fill in the settings relevant to the operations you plan to use. Each operation only reads its own settings and ignores the rest.

   **Push Missing Rows requirements:** the master must have `resident_id` and `ZoneName` columns, and each user sheet must also have a `ZoneName` column (normally pushed from master). The user sheet's assigned zone is detected automatically from the most common non-blank value in that column.

   **Pull Missing Rows requirements:** the master and each source user sheet must have a `resident_id` column. Pull operations do not use `ZoneName` filtering — any row from a user sheet with a `resident_id` absent from master is appended to master. If a user sheet has columns not yet in master, those headers are added to master before rows are appended.

   **Pull Data requirements:** the master and each source user sheet must have a `resident_id` column. Existing master rows are updated only according to the `Pull Column Policy` tab. Rows absent from master are appended, and user-sheet columns missing from master are added to master first.

6. Add rows to the Column Mapping tab — one row per column to sync:

   | Source Column | Target Column |
   |---|---|
   | column name in the source | column name in the destination |

   If both spreadsheets use the same column name, put that name in both columns. Rows with a blank cell in either column are automatically skipped.

7. Add rows to the Pull Column Policy tab for captain-entered fields:

   | Column Name | Policy |
   |---|---|
   | `Resident Name` | `fill_blank` |
   | `Damage` | `overwrite` |
   | `ZoneName` | `conflict` |
   | `resident_id` | `never` |

   Policy values:
   - `fill_blank` writes the captain value only when the master cell is blank.
   - `overwrite` replaces the master value when the captain has a non-blank value.
   - `conflict` logs differences without writing.
   - `never` skips the column entirely.

   Unlisted columns default to `conflict`, and `resident_id` is always protected as `never`.

8. For guided workflows, review the Workflow Presets tab and then click **Sheet Smart → Open Dashboard**. The first supported sidebar workflows are **Update Master From Sales Tracker** and **Push Dashboard Fields to Captain Sheets**. Start with **Dry Run** and review fills/overwrites/conflicts before running live.

   Workflow `Column Policies` use one line per target/source column:

   ```text
   Sales History -> overwrite
   Latest Sale Date -> overwrite
   ```

   Supported policies are `fill_blank`, `overwrite`, `conflict`, and `never`. Unlisted workflow columns default to `fill_blank`.

> **Note:** Phase 2 does not require the Drive Advanced Service. Only Phase 1 needs it.

### Write rules

- **Blank target cell + source has a value** → filled automatically.
- **Unchecked checkbox (boolean false) in target + source has a value** → treated as blank and filled; checkbox format is removed from that cell.
- **Target cell already has a value that differs from source** → logged as a **Conflict**, left unchanged. You decide what to do.
- **Missing column in target sheet** → column header is appended to the end of the header row; all new cells start as plain General format (no checkboxes, no false values).
- **Push Missing Rows** appends whole new rows at the bottom of the target, populated by header-name join from master (every target column with a matching master header is filled). Existing rows are never modified; rows are never removed. Running it multiple times is safe — residents already in the sheet are skipped.
- **Pull Missing Rows** appends whole new rows at the bottom of the master, populated by header-name join from each user sheet. User-sheet columns missing from master are added to master first. Existing master rows are never modified; rows are never removed. Running it multiple times is safe — residents already in master are skipped.
- **Pull Data** updates existing master rows according to `Pull Column Policy`, then appends rows whose `resident_id` is not already in master. Source blank cells never overwrite master values.
- **Sidebar workflow imports** can also use per-column policies. In dry run, `overwrite` rows are listed as proposed overwrites; in live mode they are written.

### Where results go

All logs land on tabs in the **Sheet Smart Config** spreadsheet. Each live run clears and rewrites its log tab.

| Operation | Live log tab | Dry run tab |
|---|---|---|
| Import → Master | `Last Import` | `Dry Run - Import` |
| Push → User Sheet | `Last Push - User Sheet` | `Dry Run - Push User Sheet` |
| Push → User Sheets Folder | `Last Push - Folder` | `Dry Run - Push Folder` |
| Push Missing Rows → User Sheet / Folder | `Last Push - Missing Rows` | `Dry Run - Push Missing Rows` |
| Push Missing Rows (sensitive flags) | `Flagged - Sensitive Data` | `Dry Run - Flagged Sensitive Data` |
| Pull Missing Rows ← User Sheet / Folder | `Last Pull - Missing Rows` | `Dry Run - Pull Missing Rows` |
| Pull Data ← User Sheet / Folder | `Last Pull Data` | `Dry Run - Pull Data` |
| Rename Columns → User Sheets Folder | `Last Rename - Folder` | `Dry Run - Rename Folder` |
| Sidebar workflow: Update Master From Sales Tracker | `Last Run - Update Master From Sales Tracker` | `Dry Run - Update Master From Sales Tracker` |
| Sidebar workflow: Push Dashboard Fields to Captain Sheets | `Last Run - Push Dashboard Fields to Captain Sheets` | `Dry Run - Push Dashboard Fields to Captain Sheets` |

The **Flagged - Sensitive Data** tab only appears when at least one appended row had a non-blank value in one or more `Sensitive Columns`. Each row lists the destination spreadsheet, the resident_id and name, and which sensitive columns were populated — so you can go to the captain whose zone the resident came from and confirm that sharing those notes with the new captain is okay.

### Log entry types

| Type | Means |
|---|---|
| Column Added | A missing column header was appended to the target sheet's row 1; for Pull Missing Rows, this means a source-only header was added to master. |
| Filled | A blank cell was filled with the source value. |
| Conflict | The target cell already had a different value. Not overwritten — review manually. |
| Overwritten | A workflow import or Pull Data operation replaced an existing value because the column policy was `overwrite`. |
| Appended | A whole new row was added at the bottom of the target (Push Missing Rows or Pull Missing Rows only). |
| Renamed | The target header in row 1 was changed from the old name to the new name. |
| Skipped | A row or schema change was intentionally skipped (e.g. duplicate `resident_id`, blank `resident_id`, old header missing, or new header already exists). |
| Error | Something prevented the row or sheet from being processed. |