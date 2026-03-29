# Architecture

This document describes the current architecture of DatabaseManager.

See also: [Documentation Index](README.md) | [Project README](../README.md) | [User Guide](USER_GUIDE.md) | [Troubleshooting](TROUBLESHOOTING.md)

## High-Level Design

DatabaseManager is split into two projects:

- `src/DatabaseManager.Wpf`: Presentation/UI layer (WPF)
- `src/DatabaseManager.Core`: Domain models and service implementations

A separate test project validates key core services:

- `tests/DatabaseManager.Tests`

## Layer Responsibilities

## WPF Layer (`DatabaseManager.Wpf`)

Primary responsibilities:

- User interaction and UI state
- Event handling and command orchestration
- Data binding to grids and controls
- Theme switching and view behavior

Main entry points:

- `App.xaml` and `App.xaml.cs`
- `MainWindow.xaml` and `MainWindow.xaml.cs`
- Theme dictionaries:
  - `Themes/DarkTheme.xaml`
  - `Themes/LightTheme.xaml`

MainWindow orchestration includes:

- SQL execution lifecycle (run/cancel/status)
- Metadata loading and schema assistant behavior
- Output tab navigation
- Edit Rows state machine
- Stored procedure runner parameter mapping

## Core Layer (`DatabaseManager.Core`)

Primary responsibilities:

- Query execution
- Schema discovery
- Query template persistence
- Export logic
- Query assistant SQL generation
- Row-edit SQL operations
- Stored procedure execution

### Models

- Query execution/result:
  - `Models/QueryExecutionResult.cs`
- Templates:
  - `Models/QueryTemplate.cs`
- Row edit payload:
  - `Models/Editing/RowUpdateRequest.cs`
- Schema metadata:
  - `Models/Schema/*`

### Services

- Query:
  - `IDatabaseQueryService`
  - `SqlServerQueryService`
- Templates:
  - `ITemplateStoreService`
  - `TemplateStoreService`
- Export:
  - `IExportService`
  - `ExportService`
- Row edit:
  - `IRowEditService`
  - `RowEditService`
- Schema:
  - `IDatabaseSchemaService`
  - `SqlServerSchemaService`
- Query assistant:
  - `IQueryAssistantService`
  - `SqlQueryAssistantService`
- Stored procedures:
  - `IStoredProcedureExecutionService`
  - `StoredProcedureExecutionService`

## Runtime Data Flow

## SQL execution flow

1. User writes SQL in SQL Editor tab.
2. WPF invokes `IDatabaseQueryService.ExecuteAsync`.
3. Core service executes SQL with timeout and cancellation token.
4. WPF binds returned `DataTable` (if any) to Results grid.

## Schema load flow

1. User clicks Test/Refresh schema.
2. WPF calls schema service methods.
3. Tables/procedures/columns/definitions are loaded.
4. UI updates schema lists and detail panels.

## Edit Rows flow

1. User selects table and opens Edit Rows.
2. WPF computes top/filter/order settings.
3. Core loads rows using `LoadTopRowsAsync`.
4. User edits in DataGrid.
5. On save:
   - WPF computes modified rows and key snapshots.
   - Core performs PK-based updates in one transaction.
6. On delete:
   - WPF builds PK values and asks confirmation.
   - Core deletes row in one transaction.

## Stored procedure flow

1. User selects procedure.
2. WPF builds parameter editor rows.
3. WPF maps UI values to execution parameters.
4. Core executes stored procedure and returns table/output values.

## Design Notes

- Service abstractions are interface-driven, keeping WPF decoupled from concrete implementations.
- Most app logic is event-driven in `MainWindow`.
- Query templates persist to a JSON file under local app data.
- Edit row SQL protects values through parameters and validates primary keys before write/delete.

## Testing Strategy

Current tests focus on deterministic core services:

- `ExportService`
- `SqlQueryAssistantService`
- `TemplateStoreService`

Potential future tests:

- Row edit service update/delete SQL paths
- Schema service fallback behavior
- Stored procedure execution edge cases

## Extension Points

Potential evolution paths:

- Add provider abstraction for non-SQL Server engines
- Move MainWindow orchestration to MVVM command/view-model structure
- Add optimistic concurrency support in row edit operations
- Expand test coverage for database-integrated scenarios
