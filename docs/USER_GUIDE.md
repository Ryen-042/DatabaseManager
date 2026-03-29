# User Guide

This guide explains the main workflows in DatabaseManager.

See also: [Documentation Index](README.md) | [Project README](../README.md) | [Architecture](ARCHITECTURE.md) | [Troubleshooting](TROUBLESHOOTING.md)

## Main UI Regions

- Left: Schema Assistant panel
- Center/Right: Output tab panel (`SQL Editor`, `Schema`, `Results`, `Edit Rows`, `Procedure Runner`)
- Bottom: Connection and execution controls
- Footer: Status messages

## Connect to SQL Server

1. Enter a valid connection string.
2. Set timeout (seconds).
3. Click **Test**.

If successful, schema metadata loads automatically.

## Work with Schema Assistant

## Tables tab

- Refresh schema list
- Search/filter table names
- Select table to load columns and generated table schema text

Context menu actions:

- Generate SELECT
- Generate INSERT
- Generate UPDATE
- Generate DELETE
- Get Table SQL Schema
- Copy Name

## Stored Procedures tab

- Refresh schema list
- Search/filter procedures
- Select procedure to load parameters and definition

Context menu actions:

- Generate EXEC
- Open In Runner
- Copy Name

## Query Templates tab

- Save current SQL using a template name
- Double-click template to load into SQL Editor
- Refresh/Delete templates

## SQL Editor Workflow

1. Open **SQL Editor** tab.
2. Type SQL manually or generate SQL from context menu actions in Schema Assistant.
3. Click **Run**.
4. Review tabular output in **Results**.

Keyboard shortcuts:

- `Ctrl+E`: Run query
- `Ctrl+Q`: Cancel execution

## Results Workflow

- Review returned table data.
- Export current results:
  - CSV
  - Excel

## Edit Rows Workflow

Edit Rows allows targeted table-data editing with load criteria.

## Step-by-step

1. Select a table from Schema Assistant.
2. Open **Edit Rows**.
3. Set inputs:
   - `Top` (required)
   - `Filter` (optional SQL predicate body)
   - `Order By` (optional SQL ordering expression)
4. Confirm generated SQL in the editable query box.
5. Click **Load**.
6. Edit cell values directly.
7. Click **Save Changes** to persist updates.
8. Use right-click **Delete Row...** for row deletion.
9. Click **Discard Changes** to reject unsaved grid edits.

## Important notes

- Save/Delete requires primary key columns.
- Save/Delete operations are transactional.
- Input fields and generated query are synchronized both ways.

## Procedure Runner Workflow

1. Select stored procedure in Schema Assistant.
2. Open **Procedure Runner**.
3. Enter parameter values.
4. Mark NULL as needed.
5. Click **Execute Procedure**.

Output appears in Results handling path with status updates.

## Theme and Layout Controls

- Dark mode checkbox toggles app theme.
- Object Explorer checkbox shows/hides Schema Assistant panel.

## Status Messages

The footer status bar reports:

- Connection/test outcomes
- Execution progress
- Save/delete outcomes
- Validation messages (missing selection, invalid top value, etc.)

## Practical Tips

- Keep timeout reasonable for large queries.
- Start with narrower `Top` values in Edit Rows for responsiveness.
- Use templates for frequently used scripts.
- Use query generation as a starting point and refine before running writes.
