using System.Collections.ObjectModel;
using System.Windows;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using DatabaseManager.Core.Models.Schema;
using DatabaseManager.Core.Services.Schema;
using DatabaseManager.Wpf.SqlSuggestions;

namespace DatabaseManager.Wpf.ViewModels;

/// <summary>
/// Backs the Tables/Procedures sidebar tabs and the Schema detail output tab. Selecting a
/// table or procedure here has effects on other tabs (Edit Rows, SQL Editor, Procedure Runner)
/// that this class doesn't own - those are reported via the constructor-supplied callbacks,
/// the same callback-injection pattern TemplatesPanelViewModel uses, just with more hooks
/// because this tab's selection drives more of the app.
/// </summary>
public sealed partial class SchemaAssistantViewModel : ObservableObject
{
    private readonly IDatabaseSchemaService _schemaService;
    private readonly IQueryAssistantService _queryAssistantService;
    private readonly ISqlCompletionCatalogService _completionCatalogService;
    private readonly Func<string?> _getConnectionString;
    private readonly Action<TableSchemaInfo, List<ColumnSchemaInfo>> _onTableSelected;
    private readonly Action _onTableCleared;
    private readonly Action<StoredProcedureSchemaInfo, List<StoredProcedureParameterInfo>> _onProcedureSelected;
    private readonly Action _onProcedureCleared;
    private readonly Action<List<TableSchemaInfo>, List<StoredProcedureSchemaInfo>, List<ForeignKeySchemaInfo>> _onSchemaMetadataLoaded;
    private readonly Action<string, string> _onScriptGenerated;
    private readonly Action<StoredProcedureSchemaInfo, List<StoredProcedureParameterInfo>> _onOpenInRunnerRequested;
    private readonly Func<string, Task> _onCopyRequested;
    private readonly Action<string> _setStatus;

    private List<TableSchemaInfo> _allTables = new();
    private List<StoredProcedureSchemaInfo> _allStoredProcedures = new();

    public SchemaAssistantViewModel(
        IDatabaseSchemaService schemaService,
        IQueryAssistantService queryAssistantService,
        ISqlCompletionCatalogService completionCatalogService,
        Func<string?> getConnectionString,
        Action<TableSchemaInfo, List<ColumnSchemaInfo>> onTableSelected,
        Action onTableCleared,
        Action<StoredProcedureSchemaInfo, List<StoredProcedureParameterInfo>> onProcedureSelected,
        Action onProcedureCleared,
        Action<List<TableSchemaInfo>, List<StoredProcedureSchemaInfo>, List<ForeignKeySchemaInfo>> onSchemaMetadataLoaded,
        Action<string, string> onScriptGenerated,
        Action<StoredProcedureSchemaInfo, List<StoredProcedureParameterInfo>> onOpenInRunnerRequested,
        Func<string, Task> onCopyRequested,
        Action<string> setStatus)
    {
        _schemaService = schemaService;
        _queryAssistantService = queryAssistantService;
        _completionCatalogService = completionCatalogService;
        _getConnectionString = getConnectionString;
        _onTableSelected = onTableSelected;
        _onTableCleared = onTableCleared;
        _onProcedureSelected = onProcedureSelected;
        _onProcedureCleared = onProcedureCleared;
        _onSchemaMetadataLoaded = onSchemaMetadataLoaded;
        _onScriptGenerated = onScriptGenerated;
        _onOpenInRunnerRequested = onOpenInRunnerRequested;
        _onCopyRequested = onCopyRequested;
        _setStatus = setStatus;
    }

    public ObservableCollection<TableItemViewModel> FilteredTables { get; } = new();
    public ObservableCollection<StoredProcedureItemViewModel> FilteredStoredProcedures { get; } = new();

    public IReadOnlyList<ForeignKeySchemaInfo> ForeignKeys { get; private set; } = new List<ForeignKeySchemaInfo>();

    [ObservableProperty]
    private string _tableSearchText = string.Empty;

    [ObservableProperty]
    private string _procedureSearchText = string.Empty;

    [ObservableProperty]
    private TableItemViewModel? _selectedTableItem;

    [ObservableProperty]
    private StoredProcedureItemViewModel? _selectedProcedureItem;

    [ObservableProperty]
    private List<ColumnSchemaInfo> _selectedColumns = new();

    [ObservableProperty]
    private List<StoredProcedureParameterInfo> _selectedProcedureParameters = new();

    [ObservableProperty]
    private string _tableSqlDefinitionText = string.Empty;

    [ObservableProperty]
    private string _procedureSqlDefinitionText = string.Empty;

    [ObservableProperty]
    private string _selectedTableSummary = "Select a table to inspect columns.";

    [ObservableProperty]
    private string _selectedProcedureSummary = "Select a stored procedure to inspect parameters.";

    [ObservableProperty]
    private string _schemaSummaryText = "Select a table or stored procedure from Schema Assistant.";

    [ObservableProperty]
    private int _schemaAssistantTabIndex;

    [ObservableProperty]
    private Visibility _tableDetailsVisibility = Visibility.Visible;

    [ObservableProperty]
    private Visibility _procedureDetailsVisibility = Visibility.Collapsed;

    public async Task LoadAsync(string connectionString)
    {
        _setStatus("Loading schema metadata...");

        try
        {
            var tablesTask = _schemaService.GetTablesAsync(connectionString, CancellationToken.None);
            var proceduresTask = _schemaService.GetStoredProceduresAsync(connectionString, CancellationToken.None);
            var foreignKeysTask = _schemaService.GetForeignKeysAsync(connectionString, CancellationToken.None);

            await Task.WhenAll(tablesTask, proceduresTask, foreignKeysTask);

            var tables = tablesTask.Result.ToList();
            var storedProcedures = proceduresTask.Result.ToList();
            var foreignKeys = foreignKeysTask.Result.ToList();

            _allTables = tables;
            _allStoredProcedures = storedProcedures;
            ForeignKeys = foreignKeys;
            _completionCatalogService.RefreshSchemaMetadata(tables, storedProcedures, foreignKeys);
            _onSchemaMetadataLoaded(tables, storedProcedures, foreignKeys);

            SelectedTableItem = null;
            SelectedProcedureItem = null;

            ApplyTableFilter();
            ApplyProcedureFilter();

            _setStatus(tables.Count == 0 && storedProcedures.Count == 0
                ? "Schema refresh completed, but no tables or procedures were found. Verify the target database in your connection string and user permissions."
                : $"Schema loaded successfully: {tables.Count} tables, {storedProcedures.Count} procedures, {foreignKeys.Count} foreign keys.");
        }
        catch (Exception ex)
        {
            _setStatus($"Failed to load schema metadata: {ex.Message}");
        }
    }

    [RelayCommand]
    private async Task RefreshAsync()
    {
        var connectionString = _getConnectionString();
        if (string.IsNullOrWhiteSpace(connectionString))
        {
            _setStatus("Connection string is required.");
            return;
        }

        await LoadAsync(connectionString);
    }

    partial void OnTableSearchTextChanged(string value) => ApplyTableFilter();

    partial void OnProcedureSearchTextChanged(string value) => ApplyProcedureFilter();

    partial void OnSchemaAssistantTabIndexChanged(int value)
    {
        var showTableDetails = value == 0;
        var showProcedureDetails = value == 1;

        TableDetailsVisibility = showTableDetails ? Visibility.Visible : Visibility.Collapsed;
        ProcedureDetailsVisibility = showProcedureDetails ? Visibility.Visible : Visibility.Collapsed;

        if (!showTableDetails && !showProcedureDetails)
        {
            SchemaSummaryText = "Open Tables or Stored Procedures tab to display schema details.";
        }
    }

    partial void OnSelectedTableItemChanged(TableItemViewModel? value)
    {
        if (value is null)
        {
            SelectedColumns = new List<ColumnSchemaInfo>();
            TableSqlDefinitionText = string.Empty;
            SelectedTableSummary = "Select a table to inspect columns.";
            SchemaSummaryText = "Select a table or stored procedure from Schema Assistant.";
            _onTableCleared();
            return;
        }

        _ = LoadTableDetailsAsync(value.Table);
    }

    partial void OnSelectedProcedureItemChanged(StoredProcedureItemViewModel? value)
    {
        if (value is null)
        {
            SelectedProcedureParameters = new List<StoredProcedureParameterInfo>();
            ProcedureSqlDefinitionText = string.Empty;
            SelectedProcedureSummary = "Select a stored procedure to inspect parameters.";
            SchemaSummaryText = "Select a table or stored procedure from Schema Assistant.";
            _onProcedureCleared();
            return;
        }

        _ = LoadProcedureDetailsAsync(value.Procedure);
    }

    private async Task LoadTableDetailsAsync(TableSchemaInfo table)
    {
        var connectionString = _getConnectionString();
        if (string.IsNullOrWhiteSpace(connectionString))
        {
            _setStatus("Connection string is required.");
            return;
        }

        try
        {
            var columns = (await _schemaService.GetColumnsAsync(connectionString, table.SchemaName, table.TableName, CancellationToken.None))
                .OrderBy(x => x.OrdinalPosition)
                .ToList();

            SelectedColumns = columns;
            _completionCatalogService.RefreshTableColumns(table, columns);
            TableSqlDefinitionText = _queryAssistantService.BuildTableSchemaText(table, columns);
            SelectedTableSummary = $"{table.FullName} ({columns.Count} columns)";
            SchemaSummaryText = $"Table selected: {table.FullName}";
            _onTableSelected(table, columns);
        }
        catch (Exception ex)
        {
            _setStatus($"Failed to load columns: {ex.Message}");
        }
    }

    private async Task LoadProcedureDetailsAsync(StoredProcedureSchemaInfo procedure)
    {
        var connectionString = _getConnectionString();
        if (string.IsNullOrWhiteSpace(connectionString))
        {
            _setStatus("Connection string is required.");
            return;
        }

        try
        {
            var parameters = (await _schemaService.GetStoredProcedureParametersAsync(connectionString, procedure.SchemaName, procedure.ProcedureName, CancellationToken.None))
                .OrderBy(x => x.OrdinalPosition)
                .ToList();

            SelectedProcedureParameters = parameters;
            _completionCatalogService.RefreshProcedureParameters(procedure, parameters);

            var definition = await _schemaService.GetStoredProcedureDefinitionAsync(connectionString, procedure.SchemaName, procedure.ProcedureName, CancellationToken.None);
            ProcedureSqlDefinitionText = string.IsNullOrWhiteSpace(definition)
                ? "Definition is unavailable for this object or current login does not have VIEW DEFINITION permission."
                : definition;

            SelectedProcedureSummary = $"{procedure.FullName} ({parameters.Count} parameters)";
            SchemaSummaryText = $"Stored procedure selected: {procedure.FullName}";
            _onProcedureSelected(procedure, parameters);
        }
        catch (Exception ex)
        {
            _setStatus($"Failed to load procedure parameters: {ex.Message}");
        }
    }

    private void ApplyTableFilter()
    {
        var query = TableSearchText.Trim();
        var filtered = string.IsNullOrWhiteSpace(query)
            ? _allTables
            : _allTables.Where(x => x.FullName.Contains(query, StringComparison.OrdinalIgnoreCase)).ToList();

        FilteredTables.Clear();
        foreach (var table in filtered)
        {
            FilteredTables.Add(new TableItemViewModel(table));
        }
    }

    private void ApplyProcedureFilter()
    {
        var query = ProcedureSearchText.Trim();
        var filtered = string.IsNullOrWhiteSpace(query)
            ? _allStoredProcedures
            : _allStoredProcedures.Where(x => x.FullName.Contains(query, StringComparison.OrdinalIgnoreCase)).ToList();

        FilteredStoredProcedures.Clear();
        foreach (var procedure in filtered)
        {
            FilteredStoredProcedures.Add(new StoredProcedureItemViewModel(procedure));
        }
    }

    private bool EnsureTableSelection(out TableSchemaInfo table, out List<ColumnSchemaInfo> columns)
    {
        table = SelectedTableItem?.Table!;
        columns = SelectedColumns;

        if (SelectedTableItem is null)
        {
            _setStatus("Select a table first.");
            return false;
        }

        if (SelectedColumns.Count == 0)
        {
            _setStatus("No columns available for selected table.");
            return false;
        }

        return true;
    }

    private bool EnsureProcedureSelection(out StoredProcedureSchemaInfo procedure)
    {
        procedure = SelectedProcedureItem?.Procedure!;

        if (SelectedProcedureItem is null)
        {
            _setStatus("Select a stored procedure first.");
            return false;
        }

        return true;
    }

    [RelayCommand]
    private void GenerateSelect()
    {
        if (!EnsureTableSelection(out var table, out _))
        {
            return;
        }

        _onScriptGenerated(_queryAssistantService.BuildSelectTopQuery(table), "Generated SELECT query from table metadata.");
    }

    [RelayCommand]
    private void GenerateInsert()
    {
        if (!EnsureTableSelection(out var table, out var columns))
        {
            return;
        }

        _onScriptGenerated(_queryAssistantService.BuildInsertQuery(table, columns), "Generated INSERT query from table metadata.");
    }

    [RelayCommand]
    private void GenerateUpdate()
    {
        if (!EnsureTableSelection(out var table, out var columns))
        {
            return;
        }

        var sql = _queryAssistantService.BuildUpdateQuery(table, columns);
        _onScriptGenerated(sql, sql.Contains("TODO", StringComparison.Ordinal)
            ? "Generated UPDATE query. No primary key detected, so WHERE clause needs manual fix."
            : "Generated UPDATE query from table metadata.");
    }

    [RelayCommand]
    private void GenerateDelete()
    {
        if (!EnsureTableSelection(out var table, out var columns))
        {
            return;
        }

        var sql = _queryAssistantService.BuildDeleteQuery(table, columns);
        _onScriptGenerated(sql, sql.Contains("TODO", StringComparison.Ordinal)
            ? "Generated DELETE query. No primary key detected, so WHERE clause needs manual fix."
            : "Generated DELETE query from table metadata.");
    }

    [RelayCommand]
    private void GenerateTableSchema()
    {
        if (!EnsureTableSelection(out var table, out var columns))
        {
            return;
        }

        _onScriptGenerated(_queryAssistantService.BuildTableSchemaText(table, columns), "Generated table SQL schema text.");
    }

    [RelayCommand]
    private void GenerateDropTableScript()
    {
        if (!EnsureTableSelection(out var table, out _))
        {
            return;
        }

        _onScriptGenerated(_queryAssistantService.BuildDropTableScript(table), "Generated DROP TABLE script in SQL Editor.");
    }

    [RelayCommand]
    private void GenerateDropAndCreateTableScript()
    {
        if (!EnsureTableSelection(out var table, out var columns))
        {
            return;
        }

        _onScriptGenerated(_queryAssistantService.BuildDropAndCreateTableScript(table, columns), "Generated DROP + CREATE TABLE script in SQL Editor.");
    }

    [RelayCommand]
    private void GenerateExec()
    {
        if (!EnsureProcedureSelection(out var procedure))
        {
            return;
        }

        _onScriptGenerated(_queryAssistantService.BuildExecuteProcedureQuery(procedure, SelectedProcedureParameters), "Generated EXEC query from stored procedure metadata.");
    }

    [RelayCommand]
    private void GenerateDropProcedureScript()
    {
        if (!EnsureProcedureSelection(out var procedure))
        {
            return;
        }

        _onScriptGenerated(_queryAssistantService.BuildDropProcedureScript(procedure), "Generated DROP PROCEDURE script in SQL Editor.");
    }

    [RelayCommand]
    private void GenerateAlterProcedureScript()
    {
        if (!EnsureProcedureSelection(out var procedure))
        {
            return;
        }

        var definition = string.IsNullOrWhiteSpace(ProcedureSqlDefinitionText) ? null : ProcedureSqlDefinitionText;
        _onScriptGenerated(_queryAssistantService.BuildAlterProcedureScript(procedure, definition), "Generated ALTER PROCEDURE script in SQL Editor.");
    }

    [RelayCommand]
    private void OpenInRunner()
    {
        if (!EnsureProcedureSelection(out var procedure))
        {
            return;
        }

        _onOpenInRunnerRequested(procedure, SelectedProcedureParameters);
    }

    [RelayCommand]
    private async Task CopySelectedObjectNameAsync()
    {
        var selectedName = SelectedTableItem?.FullName ?? SelectedProcedureItem?.FullName ?? string.Empty;
        if (string.IsNullOrWhiteSpace(selectedName))
        {
            _setStatus("Select a table or stored procedure first.");
            return;
        }

        await _onCopyRequested(selectedName);
    }
}
