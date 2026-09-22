# Architecture

This document describes the current architecture of DatabaseManager.

See also: [Documentation Index](README.md) | [Project README](../README.md) | [User Guide](USER_GUIDE.md) | [Troubleshooting](TROUBLESHOOTING.md)

## High-Level Design

Three projects under one solution (`DatabaseManager.slnx`):

- `src/DatabaseManager.Core` (`net8.0`): domain models and service implementations, no UI dependency. Interface-driven (`I*Service` + concrete implementation) so the WPF layer never talks to ADO.NET directly.
- `src/DatabaseManager.Wpf` (`net8.0-windows`, WinExe): presentation layer.
- `tests/DatabaseManager.Tests` (xUnit): covers pure/deterministic logic in both projects — see [Testing Strategy](#testing-strategy).

## Layer Responsibilities

### WPF Layer (`DatabaseManager.Wpf`)

MVVM migration is incremental and ongoing, tab by tab, rather than a big-bang rewrite. As of this writing:

- **Fully extracted to ViewModel + View**: Query Templates, Schema Assistant (Tables/Procedures/Schema-detail), Procedure Runner.
- **Partially extracted**: the SQL Editor tab — `QueryDocumentViewModel` owns Run/Cancel execution state and commands, but the AvalonEdit control itself and the SQL suggestion popup stay in `MainWindow.xaml.cs` (see below).
- **Deliberately not extracted**: Edit Rows. Its grid binds to a live `DataTable.DefaultView`, and row-added/modified/deleted coloring is applied by walking `DataGridRow` containers directly off each row's `RowState`. Moving that to a ViewModel would mean redesigning the row model itself (wrapping every row in a bindable object), not just relocating methods — nobody has decided that redesign is worth its risk, since this is also the code that saves and deletes real data. Only the *pure* logic underneath Edit Rows (SQL-text construction, the no-PK delete safeguard, structured⟷SQL-text sync) has been pulled into tested, dependency-free Core classes; the DataGrid orchestration around them stays in `MainWindow.xaml.cs`.
- **Not built yet**: multi-tab query documents. There's a single SQL Editor/Results pair today, not a document strip.

Folder guide:

- `Commands/` — `AppCommandRegistry` (`ICommandRegistry`) is the single source of truth for every app-level action: id, category, icon key, optional keyboard gesture, and the backing `ICommand`. The top menu, the command palette (Ctrl+Shift+P), the shortcuts help panel (Ctrl+/), and `Window.InputBindings` are all built from this one registry, so a shortcut can't drift out of sync between them.
- `ViewModels/` — one ViewModel per extracted tab (see above), plus `MainWindowViewModel` as the composition root exposing them. Extracted ViewModels take callbacks from `MainWindow` for anything outside their own tab's concern (switching output tabs, clipboard access, status text, rendering shared Results) rather than holding a reference to `MainWindow` or any XAML element — this keeps them independently constructible and testable with fakes.
- `Views/` — a XAML `UserControl` per extracted tab, bound to its ViewModel via a property on `MainWindowViewModel`.
- `Windows/` — standalone dialogs (`ConnectionPickerWindow` — a single unified view (not tabs): saved connections with a dimmed/ellipsis-trimmed connection-string preview, a raw-connection-string fallback box, and a Databases section that auto-fetches `sys.databases` on whichever of the two is currently active and rewrites `Initial Catalog` via `ConnectionStringHelper` when a database is picked, `CommandPaletteWindow`, `ShortcutsHelpWindow`, `QueryParametersWindow`, `DeleteColumnsSelectionWindow`), each with a small constructor-in/result-property-out contract.
- `Controls/` — `AppMenu` (generated from the command registry), `ToastHost` (renders `IToastService.Toasts`).
- `Converters/`, `Behaviors/` — small reusable XAML-facing pieces (a value converter, a DataGrid attached-property behavior) pulled out of `MainWindow.xaml.cs` once they had more than one consumer.
- `Editors/` — `ISqlTextEditor` abstraction over AvalonEdit, used by both the SQL Editor tab and the Edit Rows query box.
- `SqlSuggestions/` — the SQL autocomplete engine (schema-derived completion catalog + context-aware ranking), shared by both text editors via a single popup that still lives in `MainWindow.xaml.cs`.
- `Themes/` — dark/light brush dictionaries, hand-authored vector icons, button/spacing style tokens.

`MainWindow.xaml.cs` remains the orchestrator for everything not yet extracted: connection handling, Edit Rows state, the shared SQL suggestion popup, and Results rendering/export/marking (shared between the SQL Editor and Procedure Runner, which is why it hasn't moved into either ViewModel alone).

### Core Layer (`DatabaseManager.Core`)

Primary responsibilities: query execution, schema discovery, query template persistence, export, query-assistant SQL generation, row-edit SQL operations, stored procedure execution, saved connection profiles.

Every DB-touching service is thin I/O wrapped around **pure, independently-testable** logic where that logic is non-trivial:

| Concern | DB-coupled service | Pure logic pulled out |
|---|---|---|
| Ad-hoc SQL execution | `SqlServerQueryService` | `QueryOutputModeParser` (`-- full` directive, CREATE/ALTER routine detection), `QueryBatchSplitter` (splits on `GO` batch separators) |
| Row editing | `RowEditService` | `RowEditSqlBuilder` (UPDATE/INSERT/DELETE statement text, the no-PK delete mismatch safeguard's predicates), `RowEditQueryTextSync` (Edit Rows' structured-inputs⟷SQL-text conversion) |
| Stored procedures | `StoredProcedureExecutionService` | `ProcedureParameterMapper` (editor-row → execution-parameter mapping, NULL handling) |
| Schema-derived scripts | — | `SqlQueryAssistantService` (already pure) |

This split exists because every method on `RowEditService`/`StoredProcedureExecutionService`/`SqlServerQueryService` opens a real `SqlConnection`, so they can't be unit tested directly — but the decision logic inside them (is this a CREATE PROCEDURE? does the no-PK delete's matched row count agree with what was intended? what's the null-safe WHERE predicate?) is exactly the kind of thing worth protecting with tests before anyone touches the code that calls it.

Other services: `IDatabaseSchemaService`/`SqlServerSchemaService` (table/column/FK/procedure discovery, plus `GetDatabasesAsync` for the connection picker's Databases section), `ITemplateStoreService`/`TemplateStoreService` (query templates, JSON-persisted), `IExportService`/`ExportService` (CSV/Excel), `IConnectionProfileStoreService`/`ConnectionProfileStoreService` (saved connections, DPAPI-encrypted), `ConnectionStringHelper` (pure `SqlConnectionStringBuilder` wrapper so the WPF layer never references `Microsoft.Data.SqlClient` directly), `DisplayValueFormatter` (grid value formatting).

## Runtime Data Flow

### SQL execution (SQL Editor tab)

1. `QueryDocumentViewModel.RunCommand` reads the editor text, strips a trailing `-- full` directive (`QueryOutputModeParser.Parse`).
2. `QueryBatchSplitter.Split` breaks the remaining text on `GO` lines — most scripts are one batch, but a script copied from SSMS ("Script as CREATE") routinely has `USE`/`SET .../GO` boilerplate ahead of the statement that matters.
3. `QueryOutputModeParser.ExtractParameterNames` runs per batch and the results are unioned; a batch that's a CREATE/ALTER PROCEDURE/FUNCTION/TRIGGER contributes no parameters (its own `@param` declarations aren't query parameters to prompt for).
4. If parameters remain, `QueryParametersWindow` prompts for values.
5. `SqlServerQueryService` executes each batch sequentially on the same connection (so a `USE` in an earlier batch carries through to later ones), collecting result sets from all of them.
6. `MainWindow.DisplayExecutionResult` renders the combined result as expandable sections in the Results tab.

### Schema load flow

1. User connects or clicks Refresh in Schema Assistant.
2. `SchemaAssistantViewModel.LoadAsync` calls `IDatabaseSchemaService` for tables/procedures/foreign keys, refreshes the SQL suggestion completion catalog, and reports the loaded metadata back to `MainWindow` via callback (which keeps its own fallback-candidate fields in sync for the suggestion engine).
3. Selecting a table or procedure loads its columns/parameters/definition and reports that selection back to `MainWindow` too, which drives Edit Rows population, `TableName` placeholder substitution, and output-tab switching — side effects `SchemaAssistantViewModel` doesn't own itself.

### Edit Rows flow

1. User selects a table (via Schema Assistant) and opens Edit Rows.
2. Structured mode (Top/Filter/Order By) or custom SQL mode (the query box is authoritative) produce the SQL — `RowEditQueryTextSync` handles both directions, with `MainWindow`'s custom-mode flag deciding which one is authoritative at any moment so the other is never overwritten (this is how comments in custom SQL survive mode toggles).
3. `IRowEditService.LoadTopRowsAsync` loads rows into a `DataTable` bound directly to the grid.
4. On save, `MainWindow` computes modified/inserted rows from the `DataTable`'s row states; `RowEditSqlBuilder` builds the UPDATE/INSERT statement text; `IRowEditService.SaveRowChangesAsync` executes them in one transaction.
5. On delete: with a primary key, deletes by key in a transaction. Without one, `IRowEditService.DeleteRowsBySelectedColumnsAsync` first validates that the count of rows matching the user-selected-column predicate equals the intended selection, and refuses to proceed (offering the generated DELETE script for manual review instead) if they disagree.

### Stored procedure flow

1. "Open In Runner" (Schema Assistant) calls `ProcedureRunnerViewModel.LoadParameters`, which populates the parameter grid the runner tab binds to directly.
2. `ProcedureRunnerViewModel.ExecuteCommand` maps each grid row through `ProcedureParameterMapper` and calls `IStoredProcedureExecutionService.ExecuteAsync` — always for whichever procedure's parameters are currently loaded, not whatever happens to be selected elsewhere in Schema Assistant at click time.
3. Results render through the same `DisplayExecutionResult` path the SQL Editor uses.

## Design Notes

- Service abstractions are interface-driven, keeping WPF decoupled from concrete implementations — every Core service has an `I*Service` interface even where there's only one implementation, which is what makes the ViewModel tests below possible (fakes implement the interface).
- Grid click behavior: clicking a cell/row in the Results or Edit Rows grids does **not** copy to clipboard — that was deliberately removed because it was disruptive. Only explicit "Copy Value"/"Mark Cell"/"Mark Row" context-menu actions copy or highlight.
- Query templates persist to `%LocalAppData%/DatabaseManager/query-templates.json`; connection profiles to `%LocalAppData%/DatabaseManager/connection-profiles.json` (connection strings DPAPI-encrypted).

## Testing Strategy

`tests/DatabaseManager.Tests` covers pure/deterministic logic in both projects — 196 tests as of this writing, across:

- **Core**: the pure-logic classes listed in the table above, plus `SqlQueryAssistantService`, `TemplateStoreService`, `ExportService`, `ConnectionProfileStoreService`, `ConnectionStringHelper`.
- **WPF**: extracted ViewModels (`SchemaAssistantViewModel`, `QueryDocumentViewModel`, `QueryDocumentsViewModel`, `QueryDocumentTab`, `ProcedureRunnerViewModel`, `TemplatesPanelViewModel`, `MainWindowViewModel`, `ConnectionProfileViewModel`), each tested against a fake implementation of whichever Core service interface it depends on (no real database needed) — plus the command infrastructure (`AppCommandRegistry`, `FuzzyMatcher`, `KeyGestureFormatter`).

**No database-integration tests exist.** `RowEditService`, `StoredProcedureExecutionService`, and `SqlServerQueryService` each open a real `SqlConnection` in every public method, so their I/O plumbing itself is untested — only the decision logic factored out of them is. Treat changes to that plumbing as higher-risk and verify manually against a real connection.

## Extension Points

- Drag-to-reorder query documents — pinning already sorts documents to the front (see `QueryDocumentsViewModel`), but manual reordering within/across that isn't built.
- Redesigning Edit Rows' row model so it can move to a real ViewModel (see above) — a bigger, separate decision from the rest of the MVVM migration.
- Add provider abstraction for non-SQL-Server engines.
- Add optimistic concurrency support in row edit operations.
- A database-integration test suite (trait-gated, against a real or LocalDB SQL Server) for the services listed under Testing Strategy's caveat.
