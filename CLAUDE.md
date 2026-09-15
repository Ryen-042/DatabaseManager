# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project

DatabaseManager is a Windows desktop SQL Server utility built with WPF on .NET 8. It combines a SQL editor, schema browsing, query templates, CSV/Excel export, stored procedure execution, and transactional inline row editing in one app. SQL Server only (`Microsoft.Data.SqlClient`) — no other database provider is supported.

## Commands

```powershell
# Restore / build / run (dotnet CLI, from repo root)
dotnet restore DatabaseManager.slnx
dotnet build DatabaseManager.slnx -c Debug
dotnet run --project src/DatabaseManager.Wpf/DatabaseManager.Wpf.csproj -c Debug

# Tests (xUnit)
dotnet test DatabaseManager.slnx -c Debug
dotnet test tests/DatabaseManager.Tests/DatabaseManager.Tests.csproj          # test project only
dotnet test tests/DatabaseManager.Tests/DatabaseManager.Tests.csproj --filter "FullyQualifiedName~ClassName"   # single class
dotnet test tests/DatabaseManager.Tests/DatabaseManager.Tests.csproj --filter "FullyQualifiedName~ClassName.MethodName"  # single test

# Format
dotnet format DatabaseManager.slnx
```

A `Makefile` wraps the same flow (`make build`, `make test`, `make run`, `make clean`, `make rebuild`, `make publish`, `make publish-self-contained`, `make format`; `CONFIG=Debug|Release`) for anyone with GNU Make. Building/running requires the .NET 8 SDK and Windows (the WPF project targets `net8.0-windows`); tests build for both `net8.0` and `net8.0-windows` since the test project references the WPF assembly.

Build note: only one `DatabaseManager.Wpf.exe` process can be running at a time — a running instance locks its own `bin/` output and the next build will fail with `MSB3026`/`MSB3027` copy errors until it's closed.

## Architecture

Three projects under one solution (`DatabaseManager.slnx`):

- `src/DatabaseManager.Core` (`net8.0`) — domain models and service implementations, no UI dependency. Interface-driven (`I*Service` + concrete impl) so the WPF layer never talks to ADO.NET directly.
- `src/DatabaseManager.Wpf` (`net8.0-windows`, WinExe) — presentation layer.
- `tests/DatabaseManager.Tests` (xUnit) — covers only pure/deterministic `Core` logic; **no database-integration tests exist**. `RowEditService` and `StoredProcedureExecutionService` (the most stateful, DB-coupled services) currently have zero automated coverage — treat changes there as higher-risk and verify manually against a real connection.

### Core services (`src/DatabaseManager.Core/Services`)

| Service | Responsibility |
|---|---|
| `IDatabaseQueryService` / `SqlServerQueryService` | Ad-hoc SQL execution, multi-statement result sets |
| `IDatabaseSchemaService` / `Schema/SqlServerSchemaService` | Table/column/FK/procedure discovery |
| `IQueryAssistantService` / `Schema/SqlQueryAssistantService` | Generates SELECT/INSERT/UPDATE/DELETE/EXEC/DROP/ALTER script text |
| `IStoredProcedureExecutionService` / `Schema/StoredProcedureExecutionService` | Executes a procedure with typed/nullable params, returns results + output params |
| `IRowEditService` / `RowEditService` | Loads `TOP(N)` rows, saves updates transactionally by PK, deletes by PK or by a user-selected-column predicate (with an intended-vs-matched row-count safeguard when there's no PK) |
| `ITemplateStoreService` / `TemplateStoreService` | CRUD for saved query templates, JSON-persisted (see below) |
| `IExportService` / `ExportService` | CSV (`CsvHelper`) / Excel (`ClosedXML`) export |
| `QueryOutputModeParser` (static) | Parses a trailing `-- full` comment directive to toggle full/untruncated output, strips it before execution; also detects `CREATE`/`ALTER` `PROCEDURE`/`FUNCTION`/`TRIGGER` statements so their `@param` declarations aren't mistaken for query parameters to prompt for |
| `DisplayValueFormatter` (static) | Formats grid values (binary/array truncation unless full-output mode is on) |

### WPF layer

`MainWindow.xaml` + `MainWindow.xaml.cs` (~3000 lines) is still the orchestrator for almost everything — connection handling, schema browsing, SQL execution, Edit Rows state, stored procedure runner, and all event handlers live there as one code-behind class. There is **no MVVM in the main window yet**; `Commands/`, `ViewModels/`, `Views/` under `src/DatabaseManager.Wpf/` are scaffolding for a planned incremental migration and are currently empty except for what's listed below. Don't assume a ViewModel exists for a feature just because the folder does.

- `Windows/` — standalone dialogs that *have* been extracted to proper XAML `Window`s (`DeleteColumnsSelectionWindow`, `QueryParametersWindow`), each with a small in/out contract (constructor takes input data, exposes a result property, caller does `ShowDialog()`). New multi-step dialogs should follow this pattern rather than being built imperatively in code-behind.
- `Editors/` — `ISqlTextEditor` abstraction over AvalonEdit (`AvalonEditSqlTextEditorAdapter`, `SqlEditorSupport`), used by both the SQL Editor tab and the Edit Rows query box.
- `SqlSuggestions/` — SQL autocomplete engine: `SqlCompletionCatalogService` (schema-derived completion catalog, refreshed on schema reload) + `SqlSuggestionEngine` (context-aware ranking: keywords/snippets, tables, alias/schema-qualified columns, stored procedures, FK-based join hints, recent SQL fragments).
- `Themes/` — `DarkTheme.xaml`/`LightTheme.xaml` (brush resources only, swapped at runtime via `App.ApplyTheme(bool)` which replaces `Application.Resources.MergedDictionaries[0]` — never reorder that index), `Icons.xaml` (hand-authored vector `Geometry` resources, no color baked in), `ButtonStyles.xaml` (`PrimaryButtonStyle`/`SecondaryButtonStyle`/`DestructiveButtonStyle`/`IconButtonStyle`/`IconToggleButtonStyle` — a ghost/outline/fill-by-treatment hierarchy, not just color), `Spacing.xaml` (spacing/typography tokens). `App.xaml`'s implicit `Button`/`TabItem`/`DataGrid`/`ScrollBar` styles carry a custom re-templated look (flat, slim, underline tabs) — a `TargetType`-only style with no `x:Key` here is the *default* for every instance of that control, so check it before assuming a control needs its own explicit `Style=`.

### Data flows (span WPF + Core — read both sides before changing one)

- **SQL execution**: SQL Editor text → `QueryOutputModeParser.Parse`/`ExtractParameterNames` → (prompt via `QueryParametersWindow` if `@params` found and it's not a CREATE/ALTER routine) → `IDatabaseQueryService.ExecuteAsync` → multi-statement results rendered as expandable sections in the Results tab.
- **Edit Rows**: has two synchronized modes — *structured* (Top/Filter/Order By inputs generate the SQL) and *custom SQL* (the query box is authoritative, user comments are preserved across mode toggles and structured-input regeneration). This dual-mode sync plus the no-PK delete safeguard (blocks/warns when the DB-reported affected-row count doesn't match the intended selection) are deliberate, previously-debugged behavior — see `plan.md` Phase 3/7 for the history. Don't simplify away the comment-preservation or the mismatch check without re-reading why they're there.
- **Grid click behavior**: clicking a cell/row in the Results or Edit Rows grids does **not** copy to clipboard — that was deliberately removed (`plan.md` Phase 6) because it was disruptive. Only explicit "Copy Value"/"Mark Cell"/"Mark Row" context-menu actions copy or highlight.
- **Connection**: `App.config`'s `appSettings/DefaultConnectionString` is used for auto-connect on startup if present; the raw connection string is otherwise typed directly into the toolbar. There is no saved-profile/credential-manager layer.
- **Query templates**: persisted as JSON at `%LocalAppData%/DatabaseManager/query-templates.json` via `TemplateStoreService` — tolerates a missing file (first run) rather than erroring.

## Keyboard shortcuts

All handled imperatively in `MainWindow_PreviewKeyDown` (no declarative `KeyBinding`s) — Ctrl+1..5 switches Output tabs (Edit Rows/SQL Editor/Schema/Results/Procedure Runner, in that index order), Ctrl+R refreshes Edit Rows, Ctrl+S saves row changes (Edit Rows tab only), Ctrl+E runs the query or refreshes Edit Rows depending on the active tab, Ctrl+Q cancels execution, Ctrl+Space triggers the SQL suggestion popup (with Up/Down/Enter/Tab/Escape navigating it — that part lives in `SqlEditorTextBox_PreviewKeyDown`, separate from the global handler).

## Documentation

- `docs/ARCHITECTURE.md` — layer responsibilities and runtime data-flow diagrams (prose).
- `docs/USER_GUIDE.md`, `docs/TROUBLESHOOTING.md` — end-user facing.
- `plan.md` (repo root) — historical phased backlog of past feature work; useful for *why* a behavior exists, not a current task list.
