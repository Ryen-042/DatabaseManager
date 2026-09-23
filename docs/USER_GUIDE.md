# User Guide

This guide explains the main workflows in DatabaseManager.

See also: [Documentation Index](README.md) | [Project README](../README.md) | [Architecture](ARCHITECTURE.md) | [Troubleshooting](TROUBLESHOOTING.md)

## Main UI Regions

- Top: a menu bar (Connection / Edit Rows / Query / View), generated from the app's command
  registry — every menu item also has a matching keyboard shortcut and shows up in the command
  palette (`Ctrl+Shift+P`).
- Toolbar: the connection summary button, timeout, and Connect.
- Left: Schema Assistant panel (Tables / Procedures / Query Templates).
- Center/Right: four output tabs — **Edit Rows**, **Query** (SQL authoring + results, merged into
  one tab), **Schema**, **Procedure Runner**.
- Footer: status messages, plus toast notifications in the bottom-right for background failures.

## Connect to SQL Server

1. Click the connection summary button (shows the current connection's name, or "Not Connected")
   in the toolbar to open the connection picker.
2. In the picker you can:
   - Pick a **saved connection** from the list (its connection string previews dimmed underneath,
     trimmed if long). Manage saved connections with **New** / **Edit** / **Set Default** /
     **Delete**; connection strings are encrypted at rest (DPAPI, current-user scope).
   - Or paste a **raw connection string** directly — not saved anywhere, just used once.
   - Either way, the **Databases** list below auto-fetches that server's databases (no button
     press needed) — pick one to connect to a specific database instead of whatever the
     connection string's `Initial Catalog` already says. System databases (`master`, `model`,
     `msdb`, `tempdb`) are hidden by default; check "Show system databases" to see them.
3. Mark a saved profile as **Set Default** so it auto-connects on startup.
4. Set the timeout (seconds) in the toolbar before connecting if the default isn't right.
5. Click **Connect**. On success, schema metadata loads automatically and the connection
   summary button updates to show the active profile/database.

## Work with Schema Assistant

### Tables tab

- Refresh schema list
- Search/filter table names
- Select a table to load its columns and generated table-schema text (shown in the **Schema**
  output tab)

Right-click context menu actions:

- Generate SELECT / INSERT / UPDATE / DELETE
- Get Table SQL Schema
- **Drop Table** — generates a `DROP TABLE` script (doesn't execute it)
- **Drop + Recreate Table** — generates a script that drops and recreates the table from its
  current schema
- Copy Name

All of these insert generated SQL into the Query tab's editor without executing it — review
before running, especially the destructive ones.

### Stored Procedures tab

- Refresh schema list
- Search/filter procedures
- Select a procedure to load its parameters and definition (shown in the **Schema** output tab)

Right-click context menu actions:

- Generate EXEC
- **Open In Runner** — loads the procedure's parameters into the Procedure Runner tab
- **Generate ALTER Script** — generates an `ALTER PROCEDURE` script from the current definition
  (a safe placeholder + status message if the definition is encrypted/unavailable)
- **Drop Procedure** — generates a `DROP PROCEDURE` script
- Copy Name

### Query Templates tab

- Save the Query tab's current SQL using a template name
- Double-click a template to load it into the active query document
- Refresh/Delete templates

## Query Tab Workflow

The Query tab holds one or more **query documents** — independent SQL buffers that share one
editor, one autocomplete popup, and one results panel (only one query runs at a time across all
of them).

1. Open the **Query** tab.
2. Use the document strip above the editor to manage documents:
   - `Ctrl+T` / the **+** button: open a new document
   - `Ctrl+W` / the **×** on a tab: close the current document (prompts first if it has
     unexecuted changes — the dirty dot next to its title)
   - `Alt+1`..`Alt+9`: jump straight to the Nth open document
   - Double-click a document's title to rename it; right-click for **Rename** / **Pinned**
     (pinned documents sort to the front of the strip) / **Color** (a small palette of tab
     colors) / **Close**
3. Type SQL, or insert it from a Schema Assistant context-menu action. As you type, an
   autocomplete popup offers context-aware suggestions — trigger it manually with `Ctrl+Space`,
   navigate with `Up`/`Down`, accept with `Enter`/`Tab`, dismiss with `Escape`. It suggests:
   - keywords and snippets (`SELECT TOP`, `INSERT`/`UPDATE`/`DELETE`/`EXEC` templates, a CTE
     template)
   - table names (schema-qualified and unqualified)
   - columns, once you've typed a table alias or `schema.`
   - stored procedures and their parameters after `EXEC`
   - foreign-key-based join hints and recently-used SQL fragments
4. A script with multiple `;`-separated statements, or multiple `GO`-separated batches (e.g. a
   pasted "Script as CREATE" with a `USE`/`SET ...`/`GO` preamble), is fully supported — batches
   run sequentially on the same connection, so an earlier `USE` carries through.
5. If your SQL references `@parameters` (outside a `CREATE`/`ALTER PROCEDURE`/`FUNCTION`/
   `TRIGGER` statement, whose own `@param` declarations aren't mistaken for query parameters),
   a dialog prompts for their values before running.
6. Add a trailing `-- full` comment to a statement to see untruncated binary/array values in its
   results instead of the default preview truncation.
7. Click **Run** (or `Ctrl+E`). Cancel a running query with the **Cancel** button or `Ctrl+Q`.

### Results

Each statement's result renders as its own **expandable section** — header shows the statement
number plus row count (or affected-row count for non-`SELECT` statements) and elapsed time.
Clicking a cell or row does **not** copy anything to the clipboard (that was deliberately
removed as disruptive). Right-click a cell instead for:

- **Copy Value** — copies just that cell's value
- **Mark This Cell** / **Mark This Row** — highlights it for visual reference
- **Export This Section as CSV...** / **...as Excel...** — exports just that section's rows

## Edit Rows Workflow

Edit Rows loads a table's rows for direct, transactional inline editing. It has two synchronized
input modes:

- **Structured mode** (default): the `Top` / `Filter` / `Order By` fields generate the SQL shown
  in the query box below them.
- **Custom SQL query mode** (check the box): the query box becomes authoritative — write your own
  SQL directly. Comments in it are preserved even when toggling back to structured mode, and
  structured-mode regeneration never silently overwrites custom SQL/comments.

### Step-by-step

1. Select a table from Schema Assistant (or write custom SQL directly in Custom SQL query mode).
2. Open **Edit Rows**.
3. In structured mode, set:
   - `Top` (row limit)
   - `Filter` (optional `WHERE` predicate body)
   - `Order By` (optional `ORDER BY` expression)
4. Click **Load** (`Ctrl+R` also reloads while this tab is active).
5. Edit cell values directly in the grid — added/modified/deleted rows are color-coded.
6. Click **Save Changes** (`Ctrl+S`) to commit updates in a SQL transaction. Requires a primary
   key on the table.
7. Click **Discard Changes** to reject unsaved grid edits.
8. Right-click a cell for **Copy Value** if you need it on the clipboard — clicking alone does
   not copy, same as the Results grid.

### Deleting rows

- Right-click a row and choose **Delete Row...**.
- **If the table has a primary key**, rows are deleted directly by PK in a transaction.
- **If it doesn't**, a dialog lists every column as a checkbox so you can build a predicate from
  the columns you pick (at least one required) instead. Before deleting, the app runs a
  validation count query with that predicate and compares the matched row count against how many
  rows you actually selected:
  - **Counts match**: proceeds after a final confirmation.
  - **Counts don't match** (the predicate would hit more or fewer rows than intended — e.g. two
    rows share the same values in every selected column): the delete is **blocked**, and you're
    offered **Copy Generated DELETE SQL** to review and run it manually instead of risking an
    unintended row count.

## Procedure Runner Workflow

1. Select a stored procedure in Schema Assistant and choose **Open In Runner** from its context
   menu (this loads its parameters here — whichever procedure is loaded in the runner is what
   executes, regardless of what's currently selected back in Schema Assistant).
2. Open **Procedure Runner**.
3. Enter parameter values; check "Send as NULL" for a parameter instead of typing a value where
   applicable.
4. Click **Execute Procedure**.

Results (and any output parameters) render the same expandable-section way as the Query tab's
results.

## Productivity Features

- **Command Palette** (`Ctrl+Shift+P`, or View menu → Command Palette): fuzzy-searchable list of
  every app command, each showing its shortcut.
- **Shortcuts Help** (`Ctrl+/`): a searchable panel listing every shortcut, grouped by category —
  the authoritative reference if this guide's list drifts.
- **Toast notifications**: transient bottom-right notifications for background failures
  (connection, query/procedure execution), alongside the status bar text.
- **Dark/light theme** and **Schema Assistant panel visibility** toggle from the View menu (or
  their palette/shortcut entries) rather than a toolbar checkbox.

## Status Messages

The footer status bar reports:

- Connection/test outcomes
- Execution progress
- Save/delete outcomes
- Validation messages (missing selection, invalid `Top` value, delete mismatch, etc.)

## Keyboard Shortcuts

- `Ctrl+1`..`Ctrl+4`: switch output tabs (Edit Rows / Query / Schema / Procedure Runner)
- `Ctrl+E`: run the query, or refresh Edit Rows if that tab is active
- `Ctrl+Q`: cancel a running query
- `Ctrl+T` / `Ctrl+W`: new / close query document
- `Alt+1`..`Alt+9`: switch to the Nth open query document
- `Ctrl+R`: refresh Edit Rows
- `Ctrl+S`: save Edit Rows changes
- `Ctrl+Space`: trigger SQL suggestions (`Up`/`Down`/`Enter`/`Tab`/`Escape` to navigate)
- `Ctrl+Shift+P`: open the command palette
- `Ctrl+/`: open the shortcuts help panel

The palette and help panel always reflect the full, current set — check there if this list drifts.

## Practical Tips

- Keep timeout reasonable for large queries.
- Start with narrower `Top` values in Edit Rows for responsiveness.
- Use templates for frequently used scripts.
- Use query generation as a starting point and refine before running writes — generated scripts
  are inserted, never auto-executed.
- Pin and color query documents you're keeping open across a long session so they don't get lost
  in the strip.
