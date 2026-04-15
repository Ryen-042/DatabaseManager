# DatabaseManager

A desktop SQL utility for SQL Server built with WPF on .NET 8.

DatabaseManager combines query authoring, schema exploration, template management, result export, stored procedure execution, and transactional row editing in a single Windows app.

## Highlights

- SQL editor tab with multiline query authoring.
- Schema Assistant panel for tables, procedures, and query templates.
- Query generation helpers (SELECT, INSERT, UPDATE, DELETE, EXEC, and table schema script).
- Query template persistence in local app data.
- Result viewing and export to CSV and Excel.
- Stored procedure runner with parameter input.
- Edit Rows mode:
  - Load `TOP (N)` rows from a selected table.
  - Optional `WHERE` predicate and `ORDER BY` expression.
  - Two-way sync between filter inputs and generated editable SQL.
  - PK-based update and delete operations in SQL transactions.
- Dark and light themes.
- Object Explorer toggle and keyboard shortcuts.

## Screenshots

Save the screenshots under `docs/images/` using the filenames below, then they will render here automatically.

### Edit Rows Tab

![Edit Rows Tab](docs/images/edit-rows-tab.png)

### Schema Tab

![Schema Tab](docs/images/schema-tab.png)

## Tech Stack

- UI: WPF (`net8.0-windows`)
- Core logic/services: .NET (`net8.0`)
- SQL provider: `Microsoft.Data.SqlClient`
- Exports:
  - CSV: `CsvHelper`
  - Excel: `ClosedXML`
- Tests: xUnit

## Solution Layout

```text
DatabaseManager.slnx
src/
  DatabaseManager.Core/
  DatabaseManager.Wpf/
tests/
  DatabaseManager.Tests/
```

See detailed architecture in [docs/ARCHITECTURE.md](docs/ARCHITECTURE.md).

## Prerequisites

- Windows 10/11
- .NET 8 SDK
- Access to a SQL Server instance

Optional:
- GNU Make, if you want to use the provided `Makefile` shortcuts

## Quick Start

### Using dotnet CLI

```powershell
dotnet restore DatabaseManager.slnx
dotnet build DatabaseManager.slnx -c Debug
dotnet run --project src/DatabaseManager.Wpf/DatabaseManager.Wpf.csproj -c Debug
```

### Run tests

```powershell
dotnet test DatabaseManager.slnx -c Debug
```

### Using Makefile (optional)

```powershell
make build CONFIG=Debug
make run CONFIG=Debug
make test CONFIG=Debug
```

## How To Use

### 1. Connect and load metadata

1. Enter a SQL Server connection string in the bottom connection panel.
2. Set timeout in seconds.
3. Click **Test** to validate and load schema metadata.

### 2. Browse schema

Use the left **Schema Assistant** tabs:

- **Tables**: search and inspect table columns/definition.
- **Stored Procedures**: search and inspect parameters/definition.
- **Query Templates**: save, load, refresh, and delete templates.

### 3. Author and run SQL

1. Open the **SQL Editor** tab.
2. Enter SQL or use generation actions from object context menus.
3. Click **Run**.
4. Review output in **Results**.

### 4. Edit table rows

1. Select a table from Schema Assistant.
2. Open **Edit Rows** tab.
3. Configure:
   - `Top`
   - `Filter` (optional SQL predicate body)
   - `Order By` (optional SQL order expression)
4. Click **Load**.
5. Modify rows directly in grid.
6. Click **Save Changes** to commit updates.
7. Right-click row and choose **Delete Row...** to remove a row.

Notes:

- Save and delete require a primary key.
- Save and delete actions are transactional.

### 5. Execute stored procedures

1. Select procedure in Schema Assistant.
2. Open **Procedure Runner**.
3. Provide parameter values.
4. Click **Execute Procedure**.

### 6. Export results

In **Results** tab:

- **Export CSV**
- **Export Excel**

## Keyboard Shortcuts

- `Ctrl+E`: Run query
- `Ctrl+Q`: Cancel running query

## Data and Storage

- Query templates are stored in:
  - `%LocalAppData%/DatabaseManager/query-templates.json`

## Security and Safety Notes

- Protect credentials in connection strings.
- `Filter` and `Order By` in Edit Rows are interpreted as SQL fragments for row loading; use trusted input.
- Row value writes are parameterized.
- SQL object names are escaped in row-edit operations.

## Known Limitations

- Focused on SQL Server (`Microsoft.Data.SqlClient`).
- Edit Rows update/delete requires table primary key metadata.
- Query execution is single batch per run button action.

## Documentation

- Documentation Index: [docs/README.md](docs/README.md)
- Architecture: [docs/ARCHITECTURE.md](docs/ARCHITECTURE.md)
- User Guide: [docs/USER_GUIDE.md](docs/USER_GUIDE.md)
- Troubleshooting: [docs/TROUBLESHOOTING.md](docs/TROUBLESHOOTING.md)

## Testing Coverage

Current tests include:

- Export service behavior
- Query assistant SQL generation
- Template storage CRUD behavior

See: [tests/DatabaseManager.Tests](tests/DatabaseManager.Tests)

## Contributing

1. Create a branch.
2. Build and run tests.
3. Keep UI and service behavior aligned with existing architecture.
4. Submit a PR with clear change notes and validation steps.
