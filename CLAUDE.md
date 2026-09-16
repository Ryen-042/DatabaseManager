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
| `IConnectionProfileStoreService` / `ConnectionProfileStoreService` | CRUD for saved connection profiles, JSON-persisted with DPAPI-encrypted connection strings (`[SupportedOSPlatform("windows")]`) |

### WPF layer

`MainWindow.xaml` + `MainWindow.xaml.cs` (~2850 lines) is still the orchestrator for most of the app — connection handling, SQL execution, Edit Rows state, stored procedure runner, and most event handlers live there as one code-behind class. MVVM migration is incremental and in progress, tab by tab; Templates and the Schema Assistant (Tables/Procedures/Schema detail) are fully extracted so far (see `ViewModels/`/`Views/` below) — don't assume a ViewModel/View exists for any other tab just because the folders do.

- `Commands/` — `AppCommandRegistry` (`ICommandRegistry`): the single source of truth for every app-level action (id, category, icon key, optional `KeyGesture`, the backing `ICommand`). `AppCommandDescriptor` is the record type; `FuzzyMatcher` scores palette search matches; `KeyGestureFormatter` renders a `KeyGesture` for display (handles WPF's `Key.D0`-`D9`/`OemQuestion`-etc. naming so shortcuts show as "Ctrl+1"/"Ctrl+/" instead of raw enum names). `MainWindow.RegisterCommands()` registers everything and `BuildInputBindings()` turns registered gestures into `Window.InputBindings` — this *replaced* the old imperative `MainWindow_PreviewKeyDown` switch; don't reintroduce shortcut handling there, register a command instead. Contextual behavior (e.g. Ctrl+R/S only acting while Edit Rows is the active tab) lives inside each command's execute delegate, checking `OutputTabControl.SelectedIndex` directly, the same way the old handler did.
- `ViewModels/` — `MainWindowViewModel` (owns `IsDarkMode`/`IsSchemaAssistantVisible` + exposes `CommandRegistry`, `TemplatesPanel`, and `SchemaAssistant`; MainWindow still does the actual visual work for everything not yet extracted via constructor-supplied callbacks), `ToastService`/`IToastService`/`ToastNotificationViewModel` (bottom-right auto-dismissing toasts, wired into connection and query/procedure execution failures alongside the existing status bar text), `TemplatesPanelViewModel`/`TemplateItemViewModel` (Templates tab: Save/Refresh/Delete commands against `ITemplateStoreService`, plus an `ActivateSelected()` method for double-click-to-load), `SchemaAssistantViewModel`/`TableItemViewModel`/`StoredProcedureItemViewModel` (Tables/Procedures sidebar tabs + the Schema detail output tab: list+filter, selection-driven column/parameter/definition loading, and all the "Generate SELECT/INSERT/.../Open In Runner/Copy Name" commands — the biggest single-VM extraction so far since selecting a table or procedure here has side effects on Edit Rows, the SQL Editor, and the Procedure Runner that this VM doesn't own itself). All of these take callbacks from `MainWindow` for anything that reaches outside their own tab (applying a loaded template, substituting the `TableName` placeholder, switching `OutputTabControl`, populating the Procedure Runner's parameter grid, clipboard access, status text) — none of the extracted ViewModels holds a direct reference to `MainWindow` or any XAML element. `MainWindow` still keeps its own `_selectedTable`/`_selectedColumns`/`_selectedStoredProcedure`/`_selectedProcedureParameters`/`_tables`/`_storedProcedures`/`_foreignKeys` fields in sync via those callbacks, because Edit Rows Save/Delete, the SQL suggestion engine's fallback candidates, and Procedure Runner execution (all still code-behind, Phases 8-9) read them directly — don't remove those fields without migrating those readers too.
- `Views/` — `TemplatesPanelView` (Templates tab; code-behind is only a `MouseDoubleClick` passthrough to `TemplatesPanelViewModel.ActivateSelected()`, since `ListBox` has no built-in double-click command), `TablesPanelView`/`ProceduresPanelView` (Tables/Procedures sidebar tabs; `SelectedItem` two-way-bound to the shared `SchemaAssistantViewModel`, `ContextMenu` items bound via `PlacementTarget.DataContext.<Command>` with `RelativeSource={RelativeSource AncestorType=ContextMenu}` since a `ContextMenu` isn't in the visual tree its owner is and doesn't inherit `DataContext` automatically), `SchemaDetailView` (the "Schema" output tab; also bound to `SchemaAssistantViewModel` — one VM backs three Views here, which is normal MVVM, not a mistake). Every extracted View is a XAML `UserControl` bound to a `[Tab]ViewModel` property on `MainWindowViewModel`, wired into `MainWindow.xaml` as `<views:XView DataContext="{Binding [Tab]} />` inside the existing `TabItem` — follow this pattern for future tab extractions.
- `Controls/` — `AppMenu` (top menu, items generated entirely from `ICommandRegistry` grouped by category — never add a menu item by hand, register a command), `ToastHost` (renders `IToastService.Toasts`).
- `Windows/` — standalone dialogs extracted to proper XAML `Window`s (`DeleteColumnsSelectionWindow`, `QueryParametersWindow`, `CommandPaletteWindow` (Ctrl+Shift+P), `ShortcutsHelpWindow` (Ctrl+/), `ConnectionPickerWindow` (saved profiles + a raw-string fallback tab)), each with a small in/out contract (constructor takes input data, exposes a result property, caller does `ShowDialog()`). New multi-step dialogs should follow this pattern rather than being built imperatively in code-behind. **Gotcha**: if a modal `Window` closes itself from more than one path (e.g. a key handler *and* a `Deactivated` handler), guard against reentrancy — `Close()` triggers deactivation as part of its own sequence, so an unconditional `Close()` inside a `Deactivated` handler will double-close and crash (see `CommandPaletteWindow.RequestClose()` for the pattern: a single guarded close method every path calls through).
- `Editors/` — `ISqlTextEditor` abstraction over AvalonEdit (`AvalonEditSqlTextEditorAdapter`, `SqlEditorSupport`), used by both the SQL Editor tab and the Edit Rows query box.
- `SqlSuggestions/` — SQL autocomplete engine: `SqlCompletionCatalogService` (schema-derived completion catalog, refreshed on schema reload) + `SqlSuggestionEngine` (context-aware ranking: keywords/snippets, tables, alias/schema-qualified columns, stored procedures, FK-based join hints, recent SQL fragments).
- `Themes/` — `DarkTheme.xaml`/`LightTheme.xaml` (brush resources only, swapped at runtime via `App.ApplyTheme(bool)` which replaces `Application.Resources.MergedDictionaries[0]` — never reorder that index), `Icons.xaml` (hand-authored vector `Geometry` resources, no color baked in), `ButtonStyles.xaml` (`PrimaryButtonStyle`/`SecondaryButtonStyle`/`DestructiveButtonStyle`/`IconButtonStyle`/`IconToggleButtonStyle` — a ghost/outline/fill-by-treatment hierarchy, not just color), `Spacing.xaml` (spacing/typography tokens). `App.xaml`'s implicit `Button`/`TabItem`/`DataGrid`/`ScrollBar` styles carry a custom re-templated look (flat, slim, underline tabs) — a `TargetType`-only style with no `x:Key` here is the *default* for every instance of that control, so check it before assuming a control needs its own explicit `Style=`.

### Data flows (span WPF + Core — read both sides before changing one)

- **SQL execution**: SQL Editor text → `QueryOutputModeParser.Parse`/`ExtractParameterNames` → (prompt via `QueryParametersWindow` if `@params` found and it's not a CREATE/ALTER routine) → `IDatabaseQueryService.ExecuteAsync` → multi-statement results rendered as expandable sections in the Results tab.
- **Edit Rows**: has two synchronized modes — *structured* (Top/Filter/Order By inputs generate the SQL) and *custom SQL* (the query box is authoritative, user comments are preserved across mode toggles and structured-input regeneration). This dual-mode sync plus the no-PK delete safeguard (blocks/warns when the DB-reported affected-row count doesn't match the intended selection) are deliberate, previously-debugged behavior — see `plan.md` Phase 3/7 for the history. Don't simplify away the comment-preservation or the mismatch check without re-reading why they're there.
- **Grid click behavior**: clicking a cell/row in the Results or Edit Rows grids does **not** copy to clipboard — that was deliberately removed (`plan.md` Phase 6) because it was disruptive. Only explicit "Copy Value"/"Mark Cell"/"Mark Row" context-menu actions copy or highlight.
- **Connection**: `App.config`'s `appSettings/DefaultConnectionString` is used for auto-connect on startup if present. The toolbar's `ConnectionStringTextBox` is the actual value source every connection call site reads (`TryGetConnectionString()`, etc.) but is `Visibility="Collapsed"` — the visible control is a "<name> ▾" button (`ConnectionSummaryButton`) that opens `Windows/ConnectionPickerWindow`, which either fills it from a saved, DPAPI-encrypted `ConnectionProfile` (`Core/Services/ConnectionProfileStoreService`, JSON at `%LocalAppData%/DatabaseManager/connection-profiles.json`) or from its own "Raw Connection String" tab (the old always-visible-textbox workflow, preserved as a fallback). `TouchLastUsedAsync` is called in `ConnectToDatabaseAsync`'s success branch, not when a profile is merely picked.
- **Query templates**: persisted as JSON at `%LocalAppData%/DatabaseManager/query-templates.json` via `TemplateStoreService` — tolerates a missing file (first run) rather than erroring.

## Keyboard shortcuts

Declarative, via `Window.InputBindings` built from the command registry (see `Commands/` above) — Ctrl+1..5 switches Output tabs (Edit Rows/SQL Editor/Schema/Results/Procedure Runner, in that index order), Ctrl+R refreshes Edit Rows, Ctrl+S saves row changes (Edit Rows tab only), Ctrl+E runs the query or refreshes Edit Rows depending on the active tab, Ctrl+Q cancels execution, Ctrl+Shift+P opens the command palette, Ctrl+/ opens the shortcuts help panel. Ctrl+Space still triggers the SQL suggestion popup (Up/Down/Enter/Tab/Escape navigate it) via `SqlEditorTextBox_PreviewKeyDown` directly on the AvalonEdit control — that's editor-local, not part of the command registry.

## Documentation

- `docs/ARCHITECTURE.md` — layer responsibilities and runtime data-flow diagrams (prose).
- `docs/USER_GUIDE.md`, `docs/TROUBLESHOOTING.md` — end-user facing.
- `plan.md` (repo root) — historical phased backlog of past feature work; useful for *why* a behavior exists, not a current task list.
