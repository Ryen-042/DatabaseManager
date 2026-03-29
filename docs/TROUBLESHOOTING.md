# Troubleshooting

This guide covers common issues and fixes.

See also: [Documentation Index](README.md) | [Project README](../README.md) | [User Guide](USER_GUIDE.md) | [Architecture](ARCHITECTURE.md)

## App fails to start with make error code

Symptom:

- `make: *** [run] Error -532462766`

Meaning:

- This usually indicates a .NET runtime exception occurred while launching the WPF app.

What to do:

1. Run directly to surface .NET errors:

```powershell
dotnet run --project src/DatabaseManager.Wpf/DatabaseManager.Wpf.csproj -c Debug
```

2. Build first to catch compile errors:

```powershell
dotnet build src/DatabaseManager.Wpf/DatabaseManager.Wpf.csproj -c Debug
```

3. Check Windows Application log for `.NET Runtime` events if needed.

## SQL editor not visible

Symptom:

- Main editor seems missing.

Fix:

- Open the `SQL Editor` tab in the main output tab control.

## Column headers appear without underscores in Edit Rows

Symptom:

- `ATM_ID` displayed like `ATMID`.

Cause:

- WPF interprets `_` as access-key marker in headers.

Current behavior:

- Edit Rows grid escapes underscores so headers render literally.

## Save/Delete disabled in Edit Rows

Possible causes:

- No table selected
- Not currently in edit mode
- Selected table has no primary key metadata

Actions:

1. Select table in Schema Assistant.
2. Click **Load** in Edit Rows.
3. Verify table has primary key columns.

## "No row changes detected"

Cause:

- No modified rows in current editable DataTable.

Action:

- Edit at least one cell before saving.

## Stored procedure definition unavailable

Message:

- Definition unavailable or permission issue.

Cause:

- SQL login may not have `VIEW DEFINITION` rights.

Action:

- Ask DBA to grant needed metadata permissions.

## Schema list empty after test connection

Possible causes:

- Connected to unexpected database
- User lacks metadata permissions
- No user tables/procedures in target DB

Actions:

1. Verify database in connection string.
2. Refresh schema.
3. Validate SQL login permissions.

## Template issues

Templates are stored at:

- `%LocalAppData%/DatabaseManager/query-templates.json`

If templates are missing/corrupted:

1. Close app.
2. Backup then inspect/delete JSON file.
3. Reopen app and recreate templates.

## Export issues

- Ensure Results has tabular data before export.
- Verify write permissions to selected output folder.
- If Excel export fails, ensure target file is not open.

## Build and test commands

```powershell
dotnet restore DatabaseManager.slnx
dotnet build DatabaseManager.slnx -c Debug
dotnet test DatabaseManager.slnx -c Debug
```

## When filing issues

Include:

- Exact command and error text
- Connection string pattern (without secrets)
- Steps to reproduce
- Expected vs actual behavior
- App version/commit if available
