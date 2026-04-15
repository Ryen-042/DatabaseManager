using System.Data;
using System.Collections.ObjectModel;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Interop;
using System.Windows.Input;
using System.Windows.Media;
using DatabaseManager.Core.Models.Editing;
using DatabaseManager.Core.Models;
using DatabaseManager.Core.Models.Schema;
using DatabaseManager.Core.Services;
using DatabaseManager.Core.Services.Schema;
using Microsoft.Win32;

namespace DatabaseManager.Wpf;

public partial class MainWindow : Window
{
    private const int DwmwaUseImmersiveDarkMode = 20;
    private const int DwmwaUseImmersiveDarkModeBefore20H1 = 19;
    private const int ClipboardCannotOpenHResult = unchecked((int)0x800401D0);
    private const int OutputEditRowsTabIndex = 0;
    private const int OutputSqlEditorTabIndex = 1;
    private const int OutputSchemaTabIndex = 2;
    private const int OutputResultsTabIndex = 3;
    private const int OutputProcedureRunnerTabIndex = 4;

    private readonly IDatabaseQueryService _databaseQueryService = new SqlServerQueryService();
    private readonly IRowEditService _rowEditService = new RowEditService();
    private readonly ITemplateStoreService _templateStoreService = new TemplateStoreService();
    private readonly IExportService _exportService = new ExportService();
    private readonly IDatabaseSchemaService _databaseSchemaService = new SqlServerSchemaService();
    private readonly IQueryAssistantService _queryAssistantService = new SqlQueryAssistantService();
    private readonly IStoredProcedureExecutionService _storedProcedureExecutionService = new StoredProcedureExecutionService();
    private static readonly IValueConverter ResultsValueConverter = new ResultValueConverter();

    private DataTable? _currentDataTable;
    private bool _currentFullOutputMode;
    private CancellationTokenSource? _executionCancellationTokenSource;
    private List<TableSchemaInfo> _tables = new();
    private List<StoredProcedureSchemaInfo> _storedProcedures = new();
    private List<ColumnSchemaInfo> _selectedColumns = new();
    private List<StoredProcedureParameterInfo> _selectedProcedureParameters = new();
    private TableSchemaInfo? _selectedTable;
    private StoredProcedureSchemaInfo? _selectedStoredProcedure;
    private readonly ObservableCollection<ProcedureParameterEditorRow> _runnerParameterRows = new();
    private readonly GridLength _schemaPaneExpandedWidth = new(330);
    private DataTable? _editableResultsTable;
    private bool _isEditMode;
    private bool _isSyncingEditQuery;
    private static readonly Regex EditRowsQueryRegex = new(
        @"^\s*SELECT\s+TOP\s*\(\s*(?<top>\d+)\s*\)\s+\*\s+FROM\s+(?<from>\[[^\]]+\]\.\[[^\]]+\]|\S+)(?:\s+WHERE\s+(?<where>.*?))?(?:\s+ORDER\s+BY\s+(?<order>.*?))?\s*;?\s*$",
        RegexOptions.IgnoreCase | RegexOptions.Singleline | RegexOptions.Compiled);

    public MainWindow()
    {
        InitializeComponent();
        SourceInitialized += MainWindow_SourceInitialized;
        Loaded += MainWindow_Loaded;
    }

    [DllImport("dwmapi.dll")]
    private static extern int DwmSetWindowAttribute(IntPtr hwnd, int dwAttribute, ref int pvAttribute, int cbAttribute);

    private void MainWindow_SourceInitialized(object? sender, EventArgs e)
    {
        ApplyTitleBarTheme(DarkModeCheckBox.IsChecked == true);
    }

    private async void MainWindow_Loaded(object sender, RoutedEventArgs e)
    {
        App.ApplyTheme(DarkModeCheckBox.IsChecked == true);
        ApplyTitleBarTheme(DarkModeCheckBox.IsChecked == true);
        await LoadTemplatesAsync();
        RunnerParametersDataGrid.ItemsSource = _runnerParameterRows;
        OutputTabControl.SelectedIndex = OutputEditRowsTabIndex;
        UpdateSchemaDetailsPanelByAssistantTab();
        ApplyEditModeState();
        UpdateEditQueryTextFromInputs();
        await ConnectToDatabaseAsync(triggeredOnStartup: true);
        if (_tables.Count == 0 && _storedProcedures.Count == 0)
        {
            SetStatus("Ready. Shortcuts: Ctrl+E to run query, Ctrl+Q to cancel.");
        }
    }

    private void SchemaAssistantTabControl_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (!IsLoaded)
        {
            return;
        }

        UpdateSchemaDetailsPanelByAssistantTab();
    }

    private async void MainWindow_PreviewKeyDown(object sender, KeyEventArgs e)
    {
        if (Keyboard.Modifiers != ModifierKeys.Control)
        {
            return;
        }

        if (e.Key == Key.E)
        {
            if (OutputTabControl.SelectedIndex == OutputEditRowsTabIndex)
            {
                await RefreshEditRowsAsync();
            }
            else
            {
                RunQueryButton_Click(RunQueryButton, new RoutedEventArgs());
            }

            e.Handled = true;
            return;
        }

        if (e.Key == Key.Q)
        {
            CancelButton_Click(CancelButton, new RoutedEventArgs());
            e.Handled = true;
        }
    }

    private void SchemaPaneToggleButton_Checked(object sender, RoutedEventArgs e)
    {
        SchemaAssistantColumn.Width = _schemaPaneExpandedWidth;
        SchemaAssistantSplitterColumn.Width = new GridLength(6);
        SchemaAssistantPanel.Visibility = Visibility.Visible;
        SchemaPanelSplitter.Visibility = Visibility.Visible;
    }

    private void SchemaPaneToggleButton_Unchecked(object sender, RoutedEventArgs e)
    {
        SchemaAssistantColumn.Width = new GridLength(0);
        SchemaAssistantSplitterColumn.Width = new GridLength(0);
        SchemaAssistantPanel.Visibility = Visibility.Collapsed;
        SchemaPanelSplitter.Visibility = Visibility.Collapsed;
    }

    private void DarkModeCheckBox_Checked(object sender, RoutedEventArgs e)
    {
        App.ApplyTheme(true);
        ApplyTitleBarTheme(true);
    }

    private void DarkModeCheckBox_Unchecked(object sender, RoutedEventArgs e)
    {
        App.ApplyTheme(false);
        ApplyTitleBarTheme(false);
    }

    private void ApplyTitleBarTheme(bool darkMode)
    {
        var windowHandle = new WindowInteropHelper(this).Handle;
        if (windowHandle == IntPtr.Zero)
        {
            return;
        }

        var useDark = darkMode ? 1 : 0;
        var result = DwmSetWindowAttribute(
            windowHandle,
            DwmwaUseImmersiveDarkMode,
            ref useDark,
            sizeof(int));

        if (result != 0)
        {
            DwmSetWindowAttribute(
                windowHandle,
                DwmwaUseImmersiveDarkModeBefore20H1,
                ref useDark,
                sizeof(int));
        }
    }

    private async void ConnectButton_Click(object sender, RoutedEventArgs e)
    {
        await ConnectToDatabaseAsync(triggeredOnStartup: false);
    }

    private async Task ConnectToDatabaseAsync(bool triggeredOnStartup)
    {
        if (!TryGetConnectionString(out var connectionString))
        {
            return;
        }

        SetExecutionState(true);
        SetStatus(triggeredOnStartup ? "Connecting to database..." : "Connecting to database...");

        using var cts = new CancellationTokenSource();
        var result = await _databaseQueryService.ExecuteAsync(
            connectionString,
            "SELECT 1 AS IsConnected;",
            ParseTimeoutSeconds(),
            cts.Token);

        if (result.IsSuccess)
        {
            SetStatus("Connected successfully.");
            await LoadSchemaMetadataAsync(connectionString);
        }
        else
        {
            SetStatus($"Connection failed: {result.ErrorMessage}");
        }

        SetExecutionState(false);
    }

    private async void RunQueryButton_Click(object sender, RoutedEventArgs e)
    {
        if (!TryGetConnectionString(out var connectionString))
        {
            return;
        }

        var sql = QueryTextBox.Text;

        if (string.IsNullOrWhiteSpace(sql))
        {
            SetStatus("SQL query is required.");
            return;
        }

        var outputMode = QueryOutputModeParser.Parse(sql);
        var sqlToExecute = outputMode.Sql;
        if (string.IsNullOrWhiteSpace(sqlToExecute))
        {
            SetStatus("SQL query is required.");
            return;
        }

        var fullOutputEnabled = outputMode.HasFullDirective || FullOutputCheckBox.IsChecked == true;
        _currentFullOutputMode = fullOutputEnabled;

        var parameterNames = QueryOutputModeParser.ExtractParameterNames(sqlToExecute);
        IReadOnlyList<QueryParameterValue> queryParameters = Array.Empty<QueryParameterValue>();
        if (parameterNames.Count > 0)
        {
            if (!TryPromptForQueryParameters(parameterNames, out queryParameters))
            {
                SetStatus("Query execution canceled.");
                return;
            }
        }

        _executionCancellationTokenSource?.Dispose();
        _executionCancellationTokenSource = new CancellationTokenSource();

        SetExecutionState(true);
        SetStatus(fullOutputEnabled ? "Executing query (full output mode)..." : "Executing query...");

        var result = queryParameters.Count == 0
            ? await _databaseQueryService.ExecuteAsync(
                connectionString,
                sqlToExecute,
                ParseTimeoutSeconds(),
                _executionCancellationTokenSource.Token)
            : await _databaseQueryService.ExecuteAsync(
                connectionString,
                sqlToExecute,
                queryParameters,
                ParseTimeoutSeconds(),
                _executionCancellationTokenSource.Token);

        DisplayExecutionResult("Query", result);
        SetExecutionState(false);
    }

    private void CancelButton_Click(object sender, RoutedEventArgs e)
    {
        _executionCancellationTokenSource?.Cancel();
        SetStatus("Cancellation requested...");
    }

    private async void SaveTemplateButton_Click(object sender, RoutedEventArgs e)
    {
        var name = TemplateNameTextBox.Text.Trim();
        var sql = QueryTextBox.Text;

        if (string.IsNullOrWhiteSpace(name))
        {
            SetStatus("Template name is required.");
            return;
        }

        if (string.IsNullOrWhiteSpace(sql))
        {
            SetStatus("Cannot save an empty SQL template.");
            return;
        }

        await _templateStoreService.SaveAsync(new QueryTemplate
        {
            Name = name,
            Sql = sql
        }, CancellationToken.None);

        await LoadTemplatesAsync();
        SetStatus($"Template '{name}' saved.");
    }

    private async void RefreshTemplatesButton_Click(object sender, RoutedEventArgs e)
    {
        await LoadTemplatesAsync();
        SetStatus("Templates refreshed.");
    }

    private async void DeleteTemplateButton_Click(object sender, RoutedEventArgs e)
    {
        if (TemplatesListBox.SelectedItem is not QueryTemplate selected)
        {
            SetStatus("Select a template to delete.");
            return;
        }

        var confirmation = MessageBox.Show(
            $"Delete template '{selected.Name}'?",
            "Delete Template",
            MessageBoxButton.YesNo,
            MessageBoxImage.Warning);

        if (confirmation != MessageBoxResult.Yes)
        {
            return;
        }

        await _templateStoreService.DeleteAsync(selected.Name, CancellationToken.None);
        await LoadTemplatesAsync();
        SetStatus($"Template '{selected.Name}' deleted.");
    }

    private void TemplatesListBox_MouseDoubleClick(object sender, System.Windows.Input.MouseButtonEventArgs e)
    {
        if (TemplatesListBox.SelectedItem is not QueryTemplate selected)
        {
            return;
        }

        TemplateNameTextBox.Text = selected.Name;
        QueryTextBox.Text = selected.Sql;
        OutputTabControl.SelectedIndex = OutputSqlEditorTabIndex;
        SetStatus($"Template '{selected.Name}' loaded into editor.");
    }

    private async void ExportCsvButton_Click(object sender, RoutedEventArgs e)
    {
        if (!EnsureExportableData())
        {
            return;
        }

        var dialog = new SaveFileDialog
        {
            Filter = "CSV Files (*.csv)|*.csv",
            FileName = "query-results.csv"
        };

        if (dialog.ShowDialog(this) != true)
        {
            return;
        }

        await _exportService.ExportToCsvAsync(_currentDataTable!, dialog.FileName, _currentFullOutputMode, CancellationToken.None);
        SetStatus($"CSV export complete: {dialog.FileName}");
    }

    private async void EnterEditModeButton_Click(object sender, RoutedEventArgs e)
    {
        await LoadEditableRowsAsync();
    }

    private async void ApplyEditFilterButton_Click(object sender, RoutedEventArgs e)
    {
        await LoadEditableRowsAsync();
    }

    private async void SaveRowChangesButton_Click(object sender, RoutedEventArgs e)
    {
        await SaveRowChangesAsync();
    }

    private async Task<bool> SaveRowChangesAsync()
    {
        if (!_isEditMode || _editableResultsTable is null)
        {
            return true;
        }

        if (!TryGetConnectionString(out var connectionString))
        {
            return false;
        }

        if (_selectedTable is null)
        {
            SetStatus("Select a table before saving row edits.");
            return false;
        }

        if (_selectedColumns.All(c => !c.IsPrimaryKey))
        {
            SetStatus("Cannot save edits because the selected table has no primary key.");
            return false;
        }

        var updates = BuildRowUpdates(_editableResultsTable, _selectedColumns);
        var inserts = BuildRowInserts(_editableResultsTable, _selectedColumns);
        if (updates.Count == 0 && inserts.Count == 0)
        {
            SetStatus("No row changes detected.");
            return true;
        }

        SetExecutionState(true);
        SetStatus($"Saving {updates.Count + inserts.Count} row change(s)...");

        try
        {
            var affectedRows = await _rowEditService.SaveRowChangesAsync(
                connectionString,
                _selectedTable.SchemaName,
                _selectedTable.TableName,
                _selectedColumns,
                updates,
                inserts,
                ParseTimeoutSeconds(),
                CancellationToken.None);

            _editableResultsTable.AcceptChanges();
            SetStatus($"Saved changes successfully ({affectedRows} row(s) affected).");
            return true;
        }
        catch (Exception ex)
        {
            SetStatus($"Failed to save row changes: {ex.Message}");
            return false;
        }
        finally
        {
            SetExecutionState(false);
        }
    }

    private void DiscardRowChangesButton_Click(object sender, RoutedEventArgs e)
    {
        if (!_isEditMode || _editableResultsTable is null)
        {
            return;
        }

        _editableResultsTable.RejectChanges();
        SetStatus("Discarded unsaved row changes.");
    }

    private void ResultsDataGrid_PreviewMouseRightButtonDown(object sender, MouseButtonEventArgs e)
    {
        if (sender is not DataGrid dataGrid)
        {
            return;
        }

        var row = FindAncestor<DataGridRow>(e.OriginalSource as DependencyObject);
        if (row is null)
        {
            return;
        }

        row.IsSelected = true;
        dataGrid.SelectedItem = row.Item;
    }

    private async void DeleteRowMenuItem_Click(object sender, RoutedEventArgs e)
    {
        if (!_isEditMode || _editableResultsTable is null)
        {
            return;
        }

        if (!TryGetConnectionString(out var connectionString))
        {
            return;
        }

        if (_selectedTable is null)
        {
            SetStatus("Select a table before deleting rows.");
            return;
        }

        if (_selectedColumns.All(c => !c.IsPrimaryKey))
        {
            SetStatus("Cannot delete rows because the selected table has no primary key.");
            return;
        }

        if (EditRowsDataGrid.SelectedItem is not DataRowView selectedRowView)
        {
            SetStatus("Select a row to delete.");
            return;
        }

        var keyValues = BuildPrimaryKeyValues(selectedRowView.Row, _selectedColumns);
        var keyDetails = string.Join(", ", keyValues.Select(x => $"{x.Key}={x.Value}"));

        var confirmation = MessageBox.Show(
            $"Delete this row from {_selectedTable.FullName}?{Environment.NewLine}{Environment.NewLine}{keyDetails}",
            "Delete Row",
            MessageBoxButton.YesNo,
            MessageBoxImage.Warning);

        if (confirmation != MessageBoxResult.Yes)
        {
            return;
        }

        SetExecutionState(true);
        SetStatus("Deleting selected row...");

        try
        {
            var affectedRows = await _rowEditService.DeleteRowAsync(
                connectionString,
                _selectedTable.SchemaName,
                _selectedTable.TableName,
                _selectedColumns,
                keyValues,
                ParseTimeoutSeconds(),
                CancellationToken.None);

            if (affectedRows == 0)
            {
                SetStatus("No row was deleted. It may have already changed or been removed.");
                return;
            }

            _editableResultsTable.Rows.Remove(selectedRowView.Row);
            _editableResultsTable.AcceptChanges();
            SetStatus("Row deleted successfully.");
        }
        catch (Exception ex)
        {
            SetStatus($"Failed to delete row: {ex.Message}");
        }
        finally
        {
            SetExecutionState(false);
        }
    }

    private async void ExportExcelButton_Click(object sender, RoutedEventArgs e)
    {
        if (!EnsureExportableData())
        {
            return;
        }

        var dialog = new SaveFileDialog
        {
            Filter = "Excel Files (*.xlsx)|*.xlsx",
            FileName = "query-results.xlsx"
        };

        if (dialog.ShowDialog(this) != true)
        {
            return;
        }

        await _exportService.ExportToExcelAsync(_currentDataTable!, dialog.FileName, _currentFullOutputMode, CancellationToken.None);
        SetStatus($"Excel export complete: {dialog.FileName}");
    }

    private async Task LoadTemplatesAsync()
    {
        var templates = await _templateStoreService.GetAllAsync(CancellationToken.None);
        TemplatesListBox.ItemsSource = templates;
    }

    private async void RefreshSchemaButton_Click(object sender, RoutedEventArgs e)
    {
        if (!TryGetConnectionString(out var connectionString))
        {
            return;
        }

        await LoadSchemaMetadataAsync(connectionString);
    }

    private async void TablesListBox_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
    {
        if (!TryGetConnectionString(out var connectionString))
        {
            return;
        }

        _selectedTable = TablesListBox.SelectedItem as TableSchemaInfo;
        if (_selectedTable is null)
        {
            _selectedColumns.Clear();
            TableColumnsDataGrid.ItemsSource = null;
            TableSqlDefinitionTextBox.Text = string.Empty;
            SelectedTableTextBlock.Text = "Select a table to inspect columns.";
            SchemaSummaryTextBlock.Text = "Select a table or stored procedure from Schema Assistant.";
            UpdateEditQueryTextFromInputs();
            return;
        }

        try
        {
            var columns = await _databaseSchemaService.GetColumnsAsync(
                connectionString,
                _selectedTable.SchemaName,
                _selectedTable.TableName,
                CancellationToken.None);

            _selectedColumns = columns.OrderBy(x => x.OrdinalPosition).ToList();
            TableColumnsDataGrid.ItemsSource = _selectedColumns;
            TableSqlDefinitionTextBox.Text = _queryAssistantService.BuildTableSchemaText(_selectedTable, _selectedColumns);
            SelectedTableTextBlock.Text = $"{_selectedTable.FullName} ({_selectedColumns.Count} columns)";
            SchemaSummaryTextBlock.Text = $"Table selected: {_selectedTable.FullName}";
            UpdateSchemaDetailsPanelByAssistantTab();
            UpdateEditQueryTextFromInputs();
            OutputTabControl.SelectedIndex = OutputEditRowsTabIndex;
        }
        catch (Exception ex)
        {
            SetStatus($"Failed to load columns: {ex.Message}");
        }
    }

    private async void StoredProceduresListBox_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
    {
        if (!TryGetConnectionString(out var connectionString))
        {
            return;
        }

        _selectedStoredProcedure = StoredProceduresListBox.SelectedItem as StoredProcedureSchemaInfo;
        if (_selectedStoredProcedure is null)
        {
            _selectedProcedureParameters.Clear();
            ProcedureParametersDataGrid.ItemsSource = null;
            ProcedureSqlDefinitionTextBox.Text = string.Empty;
            SelectedProcedureTextBlock.Text = "Select a stored procedure to inspect parameters.";
            SchemaSummaryTextBlock.Text = "Select a table or stored procedure from Schema Assistant.";
            return;
        }

        try
        {
            var parameters = await _databaseSchemaService.GetStoredProcedureParametersAsync(
                connectionString,
                _selectedStoredProcedure.SchemaName,
                _selectedStoredProcedure.ProcedureName,
                CancellationToken.None);

            _selectedProcedureParameters = parameters.OrderBy(x => x.OrdinalPosition).ToList();
            ProcedureParametersDataGrid.ItemsSource = _selectedProcedureParameters;
            var procedureDefinition = await _databaseSchemaService.GetStoredProcedureDefinitionAsync(
                connectionString,
                _selectedStoredProcedure.SchemaName,
                _selectedStoredProcedure.ProcedureName,
                CancellationToken.None);

            ProcedureSqlDefinitionTextBox.Text = string.IsNullOrWhiteSpace(procedureDefinition)
                ? "Definition is unavailable for this object or current login does not have VIEW DEFINITION permission."
                : procedureDefinition;

            SelectedProcedureTextBlock.Text = $"{_selectedStoredProcedure.FullName} ({_selectedProcedureParameters.Count} parameters)";
            SchemaSummaryTextBlock.Text = $"Stored procedure selected: {_selectedStoredProcedure.FullName}";
            UpdateSchemaDetailsPanelByAssistantTab();
            OutputTabControl.SelectedIndex = OutputSchemaTabIndex;
        }
        catch (Exception ex)
        {
            SetStatus($"Failed to load procedure parameters: {ex.Message}");
        }
    }

    private void TableSearchTextBox_TextChanged(object sender, System.Windows.Controls.TextChangedEventArgs e)
    {
        ApplyTableFilter();
    }

    private void ProcedureSearchTextBox_TextChanged(object sender, System.Windows.Controls.TextChangedEventArgs e)
    {
        ApplyProcedureFilter();
    }

    private void GenerateSelectButton_Click(object sender, RoutedEventArgs e)
    {
        if (!EnsureTableSelection())
        {
            return;
        }

        QueryTextBox.Text = _queryAssistantService.BuildSelectTopQuery(_selectedTable!);
        OutputTabControl.SelectedIndex = OutputSqlEditorTabIndex;
        SetStatus("Generated SELECT query from table metadata.");
    }

    private void GenerateInsertButton_Click(object sender, RoutedEventArgs e)
    {
        if (!EnsureTableSelection())
        {
            return;
        }

        QueryTextBox.Text = _queryAssistantService.BuildInsertQuery(_selectedTable!, _selectedColumns);
        OutputTabControl.SelectedIndex = OutputSqlEditorTabIndex;
        SetStatus("Generated INSERT query from table metadata.");
    }

    private void GenerateUpdateButton_Click(object sender, RoutedEventArgs e)
    {
        if (!EnsureTableSelection())
        {
            return;
        }

        var sql = _queryAssistantService.BuildUpdateQuery(_selectedTable!, _selectedColumns);
        QueryTextBox.Text = sql;
        OutputTabControl.SelectedIndex = OutputSqlEditorTabIndex;

        SetStatus(sql.Contains("TODO", StringComparison.Ordinal)
            ? "Generated UPDATE query. No primary key detected, so WHERE clause needs manual fix."
            : "Generated UPDATE query from table metadata.");
    }

    private void GenerateDeleteButton_Click(object sender, RoutedEventArgs e)
    {
        if (!EnsureTableSelection())
        {
            return;
        }

        var sql = _queryAssistantService.BuildDeleteQuery(_selectedTable!, _selectedColumns);
        QueryTextBox.Text = sql;
        OutputTabControl.SelectedIndex = OutputSqlEditorTabIndex;

        SetStatus(sql.Contains("TODO", StringComparison.Ordinal)
            ? "Generated DELETE query. No primary key detected, so WHERE clause needs manual fix."
            : "Generated DELETE query from table metadata.");
    }

    private void GenerateTableSchemaButton_Click(object sender, RoutedEventArgs e)
    {
        if (!EnsureTableSelection())
        {
            return;
        }

        QueryTextBox.Text = _queryAssistantService.BuildTableSchemaText(_selectedTable!, _selectedColumns);
        OutputTabControl.SelectedIndex = OutputSqlEditorTabIndex;
        SetStatus("Generated table SQL schema text.");
    }

    private void GenerateExecButton_Click(object sender, RoutedEventArgs e)
    {
        if (!EnsureProcedureSelection())
        {
            return;
        }

        QueryTextBox.Text = _queryAssistantService.BuildExecuteProcedureQuery(_selectedStoredProcedure!, _selectedProcedureParameters);
        OutputTabControl.SelectedIndex = OutputSqlEditorTabIndex;
        SetStatus("Generated EXEC query from stored procedure metadata.");
    }

    private void OpenRunnerButton_Click(object sender, RoutedEventArgs e)
    {
        if (!EnsureProcedureSelection())
        {
            return;
        }

        _runnerParameterRows.Clear();

        foreach (var parameter in _selectedProcedureParameters.Where(x => !x.IsReturnValue))
        {
            _runnerParameterRows.Add(new ProcedureParameterEditorRow
            {
                ParameterName = parameter.ParameterName,
                DataType = parameter.DataType,
                Value = string.Empty,
                SendAsNull = false,
                IsOutput = parameter.IsOutput
            });
        }

        RunnerProcedureTextBlock.Text = $"Ready to execute {_selectedStoredProcedure!.FullName}";
        OutputTabControl.SelectedIndex = OutputProcedureRunnerTabIndex;
        SetStatus("Loaded procedure parameters into runner.");
    }

    private async void ExecuteProcedureButton_Click(object sender, RoutedEventArgs e)
    {
        if (!TryGetConnectionString(out var connectionString))
        {
            return;
        }

        if (!EnsureProcedureSelection())
        {
            return;
        }

        _currentFullOutputMode = FullOutputCheckBox.IsChecked == true;

        SetExecutionState(true);
        SetStatus(_currentFullOutputMode ? "Executing stored procedure (full output mode)..." : "Executing stored procedure...");

        var parameters = _runnerParameterRows
            .Select(x => new StoredProcedureExecutionParameter
            {
                Name = x.ParameterName,
                Value = x.Value,
                IsOutput = x.IsOutput,
                IsInputOutput = false,
                SendAsNull = x.SendAsNull
            })
            .ToList();

        var result = await _storedProcedureExecutionService.ExecuteAsync(
            connectionString,
            _selectedStoredProcedure!.SchemaName,
            _selectedStoredProcedure.ProcedureName,
            parameters,
            ParseTimeoutSeconds(),
            CancellationToken.None);

        DisplayExecutionResult("Stored procedure", result);
        SetExecutionState(false);
    }

    private async Task LoadSchemaMetadataAsync(string connectionString)
    {
        try
        {
            SetStatus("Loading schema metadata...");

            var tablesTask = _databaseSchemaService.GetTablesAsync(connectionString, CancellationToken.None);
            var proceduresTask = _databaseSchemaService.GetStoredProceduresAsync(connectionString, CancellationToken.None);

            await Task.WhenAll(tablesTask, proceduresTask);

            _tables = tablesTask.Result.ToList();
            _storedProcedures = proceduresTask.Result.ToList();

            _selectedTable = null;
            _selectedStoredProcedure = null;
            _selectedColumns.Clear();
            _selectedProcedureParameters.Clear();
            TableColumnsDataGrid.ItemsSource = null;
            ProcedureParametersDataGrid.ItemsSource = null;
            TableSqlDefinitionTextBox.Text = string.Empty;
            ProcedureSqlDefinitionTextBox.Text = string.Empty;
            SelectedTableTextBlock.Text = "Select a table to inspect columns.";
            SelectedProcedureTextBlock.Text = "Select a stored procedure to inspect parameters.";
            SchemaSummaryTextBlock.Text = "Select a table or stored procedure from Schema Assistant.";

            ApplyTableFilter();
            ApplyProcedureFilter();

            if (_tables.Count == 0 && _storedProcedures.Count == 0)
            {
                SetStatus("Schema refresh completed, but no tables or procedures were found. Verify the target database in your connection string and user permissions.");
                return;
            }

            SetStatus($"Schema loaded successfully: {_tables.Count} tables, {_storedProcedures.Count} procedures.");
        }
        catch (Exception ex)
        {
            SetStatus($"Failed to load schema metadata: {ex.Message}");
        }
    }

    private void ApplyTableFilter()
    {
        var query = TableSearchTextBox.Text.Trim();
        var filtered = string.IsNullOrWhiteSpace(query)
            ? _tables
            : _tables
                .Where(x => x.FullName.Contains(query, StringComparison.OrdinalIgnoreCase))
                .ToList();

        TablesListBox.ItemsSource = filtered;
    }

    private void ApplyProcedureFilter()
    {
        var query = ProcedureSearchTextBox.Text.Trim();
        var filtered = string.IsNullOrWhiteSpace(query)
            ? _storedProcedures
            : _storedProcedures
                .Where(x => x.FullName.Contains(query, StringComparison.OrdinalIgnoreCase))
                .ToList();

        StoredProceduresListBox.ItemsSource = filtered;
    }

    private bool EnsureTableSelection()
    {
        if (_selectedTable is null)
        {
            SetStatus("Select a table first.");
            return false;
        }

        if (_selectedColumns.Count == 0)
        {
            SetStatus("No columns available for selected table.");
            return false;
        }

        return true;
    }

    private bool EnsureProcedureSelection()
    {
        if (_selectedStoredProcedure is null)
        {
            SetStatus("Select a stored procedure first.");
            return false;
        }

        return true;
    }

    private bool TryGetConnectionString(out string connectionString)
    {
        connectionString = ConnectionStringTextBox.Text.Trim();
        if (string.IsNullOrWhiteSpace(connectionString))
        {
            SetStatus("Connection string is required.");
            return false;
        }

        return true;
    }

    private void DisplayExecutionResult(string operationName, QueryExecutionResult result)
    {
        _isEditMode = false;
        _editableResultsTable = null;
        ApplyEditModeState();

        if (!result.IsSuccess)
        {
            SetStatus($"{operationName} failed after {result.Duration.TotalSeconds:F2}s: {result.ErrorMessage}");
            return;
        }

        _currentDataTable = result.DataTable;
        ResultsDataGrid.ItemsSource = _currentDataTable?.DefaultView;
        OutputTabControl.SelectedIndex = OutputResultsTabIndex;
        ResultsSummaryTextBlock.Text = _currentDataTable is null
            ? $"Results ({result.AffectedRows} affected rows)"
            : $"Results ({_currentDataTable.Rows.Count} rows, {_currentDataTable.Columns.Count} columns)";

        if (result.OutputParameters is { Count: > 0 })
        {
            var outputs = string.Join(", ", result.OutputParameters.Select(x => $"{x.Key}={x.Value ?? "NULL"}"));
            SetStatus($"{operationName} completed in {result.Duration.TotalSeconds:F2}s. Output: {outputs}");
            return;
        }

        SetStatus($"{operationName} completed in {result.Duration.TotalSeconds:F2}s.");
    }

    private bool EnsureExportableData()
    {
        if (_currentDataTable is not { Rows.Count: > 0 })
        {
            SetStatus("No tabular results are available to export.");
            return false;
        }

        return true;
    }

    private int ParseTimeoutSeconds()
    {
        if (int.TryParse(TimeoutTextBox.Text.Trim(), out var timeout) && timeout > 0)
        {
            return timeout;
        }

        TimeoutTextBox.Text = "30";
        return 30;
    }

    private void SetExecutionState(bool isExecuting)
    {
        RunQueryButton.IsEnabled = !isExecuting;
        ConnectButton.IsEnabled = !isExecuting;
        CancelButton.IsEnabled = isExecuting;
        SaveRowChangesButton.IsEnabled = _isEditMode && !isExecuting;
        DiscardRowChangesButton.IsEnabled = _isEditMode && !isExecuting;
        DeleteRowMenuItem.IsEnabled = _isEditMode && !isExecuting;
    }

    private void SetStatus(string message)
    {
        StatusTextBlock.Text = message;
    }

    private void SchemaDataGrid_PreviewMouseLeftButtonUp(object sender, MouseButtonEventArgs e)
    {
        if (sender is not DataGrid dataGrid)
        {
            return;
        }

        var clickedCell = FindAncestor<DataGridCell>(e.OriginalSource as DependencyObject);
        if (clickedCell is null)
        {
            return;
        }

        var copiedText = TryGetClickedCellTextForCopy(clickedCell);
        if (string.IsNullOrWhiteSpace(copiedText))
        {
            return;
        }

        if (TrySetClipboardText(copiedText))
        {
            SetStatus($"Copied value to clipboard: {TruncateForStatus(copiedText)}");
        }
        else
        {
            SetStatus("Could not access the clipboard. Please try again.");
        }
    }

    private void DataGrid_AutoGeneratingColumn(object? sender, DataGridAutoGeneratingColumnEventArgs e)
    {
        if (e.Column.Header is not string headerText || string.IsNullOrEmpty(headerText))
        {
            return;
        }

        // In WPF headers, '_' is treated as an access-key marker; doubling it renders a literal underscore.
        e.Column.Header = headerText.Replace("_", "__", StringComparison.Ordinal);
    }

    private void SchemaListBox_PreviewMouseLeftButtonUp(object sender, MouseButtonEventArgs e)
    {
        if (sender is not ListBox)
        {
            return;
        }

        var clickedItem = FindAncestor<ListBoxItem>(e.OriginalSource as DependencyObject);
        if (clickedItem is null)
        {
            return;
        }

        if (clickedItem.Content is not object item)
        {
            return;
        }

        var copiedText = item switch
        {
            TableSchemaInfo table => table.FullName,
            StoredProcedureSchemaInfo procedure => procedure.FullName,
            _ => item.ToString() ?? string.Empty
        };

        if (string.IsNullOrWhiteSpace(copiedText))
        {
            return;
        }

        if (TrySetClipboardText(copiedText))
        {
            SetStatus($"Copied value to clipboard: {TruncateForStatus(copiedText)}");
        }
        else
        {
            SetStatus("Could not access the clipboard. Please try again.");
        }
    }

    private void CopySelectedObjectNameMenuItem_Click(object sender, RoutedEventArgs e)
    {
        var selectedName = TablesListBox.SelectedItem switch
        {
            TableSchemaInfo table => table.FullName,
            _ => StoredProceduresListBox.SelectedItem is StoredProcedureSchemaInfo procedure
                ? procedure.FullName
                : string.Empty
        };

        if (string.IsNullOrWhiteSpace(selectedName))
        {
            SetStatus("Select a table or stored procedure first.");
            return;
        }

        if (TrySetClipboardText(selectedName))
        {
            SetStatus($"Copied value to clipboard: {selectedName}");
        }
        else
        {
            SetStatus("Could not access the clipboard. Please try again.");
        }
    }

    private static bool TrySetClipboardText(string text)
    {
        const int maxAttempts = 5;
        const int retryDelayMs = 35;

        for (var attempt = 1; attempt <= maxAttempts; attempt++)
        {
            try
            {
                Clipboard.SetText(text);
                return true;
            }
            catch (COMException ex) when (ex.HResult == ClipboardCannotOpenHResult && attempt < maxAttempts)
            {
                Thread.Sleep(retryDelayMs);
            }
            catch (COMException ex) when (ex.HResult == ClipboardCannotOpenHResult)
            {
                return false;
            }
        }

        return false;
    }

    private static string? TryGetClickedCellText(DataGridCell cell)
    {
        if (cell.Content is TextBlock textBlock)
        {
            return textBlock.Text;
        }

        if (cell.Content is CheckBox checkBox)
        {
            return checkBox.IsChecked?.ToString();
        }

        if (cell.Content is TextBox textBox)
        {
            return textBox.Text;
        }

        var item = cell.DataContext;
        var column = cell.Column;

        if (item is null || column is not DataGridBoundColumn boundColumn || boundColumn.Binding is not Binding binding)
        {
            return null;
        }

        var propertyPath = binding.Path?.Path;
        if (string.IsNullOrWhiteSpace(propertyPath))
        {
            return null;
        }

        if (item is DataRowView rowView && rowView.Row.Table.Columns.Contains(propertyPath))
        {
            var value = rowView[propertyPath];
            return DisplayValueFormatter.FormatForDisplay(value);
        }

        var property = item.GetType().GetProperty(propertyPath);
        if (property is null)
        {
            return null;
        }

        var propertyValue = property.GetValue(item);
        return DisplayValueFormatter.FormatForDisplay(propertyValue);
    }

    private string? TryGetClickedCellTextForCopy(DataGridCell cell)
    {
        var item = cell.DataContext;
        var column = cell.Column;

        if (item is not null && column is DataGridBoundColumn boundColumn && boundColumn.Binding is Binding binding)
        {
            var propertyPath = binding.Path?.Path;
            if (!string.IsNullOrWhiteSpace(propertyPath))
            {
                if (item is DataRowView rowView && rowView.Row.Table.Columns.Contains(propertyPath))
                {
                    var value = rowView[propertyPath];
                    return DisplayValueFormatter.FormatForDisplay(value, _currentFullOutputMode);
                }

                var property = item.GetType().GetProperty(propertyPath);
                if (property is not null)
                {
                    var propertyValue = property.GetValue(item);
                    return DisplayValueFormatter.FormatForDisplay(propertyValue, _currentFullOutputMode);
                }
            }
        }

        if (cell.Content is TextBlock textBlock)
        {
            return textBlock.Text;
        }

        if (cell.Content is CheckBox checkBox)
        {
            return checkBox.IsChecked?.ToString();
        }

        if (cell.Content is TextBox textBox)
        {
            return textBox.Text;
        }

        return null;
    }

    private void ResultsDataGrid_AutoGeneratingColumn(object? sender, DataGridAutoGeneratingColumnEventArgs e)
    {
        DataGrid_AutoGeneratingColumn(sender, e);

        if (sender == EditRowsDataGrid)
        {
            var schemaColumn = _selectedColumns.FirstOrDefault(c =>
                c.ColumnName.Equals(e.PropertyName, StringComparison.OrdinalIgnoreCase));

            if (schemaColumn is not null && (schemaColumn.IsIdentity
                || schemaColumn.DataType.Equals("rowversion", StringComparison.OrdinalIgnoreCase)
                || schemaColumn.DataType.Equals("timestamp", StringComparison.OrdinalIgnoreCase)))
            {
                e.Column.IsReadOnly = true;
            }
        }

        if (e.Column is not DataGridTextColumn textColumn)
        {
            return;
        }

        if (textColumn.Binding is not Binding binding)
        {
            return;
        }

        binding.Converter = ResultsValueConverter;
    }

    private static T? FindAncestor<T>(DependencyObject? current) where T : DependencyObject
    {
        while (current is not null)
        {
            if (current is T match)
            {
                return match;
            }

            current = VisualTreeHelper.GetParent(current);
        }

        return null;
    }

    private static string TruncateForStatus(string value)
    {
        const int maxLength = 64;
        if (value.Length <= maxLength)
        {
            return value;
        }

        return value[..maxLength] + "...";
    }

    private sealed class ResultValueConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            return DisplayValueFormatter.FormatForDisplay(value);
        }

        public object ConvertBack(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            return value;
        }
    }

    private void UpdateSchemaDetailsPanelByAssistantTab()
    {
        var selectedIndex = SchemaAssistantTabControl.SelectedIndex;
        var showTableDetails = selectedIndex == 0;
        var showProcedureDetails = selectedIndex == 1;

        TableDetailsPanel.Visibility = showTableDetails ? Visibility.Visible : Visibility.Collapsed;
        ProcedureDetailsPanel.Visibility = showProcedureDetails ? Visibility.Visible : Visibility.Collapsed;

        if (!showTableDetails && !showProcedureDetails)
        {
            SchemaSummaryTextBlock.Text = "Open Tables or Stored Procedures tab to display schema details.";
        }
    }

    private async Task RefreshEditRowsAsync()
    {
        if (!_isEditMode)
        {
            await LoadEditableRowsAsync();
            return;
        }

        if (!HasPendingEditRowsChanges())
        {
            await LoadEditableRowsAsync();
            return;
        }

        var decision = MessageBox.Show(
            "You have unsaved edits in Edit Rows.\n\nYes = Save and refresh\nNo = Discard and refresh\nCancel = Keep editing",
            "Unsaved Edit Rows Changes",
            MessageBoxButton.YesNoCancel,
            MessageBoxImage.Warning);

        if (decision == MessageBoxResult.Cancel)
        {
            SetStatus("Refresh canceled. Unsaved edits were kept.");
            return;
        }

        if (decision == MessageBoxResult.Yes)
        {
            var saveSucceeded = await SaveRowChangesAsync();
            if (!saveSucceeded)
            {
                return;
            }
        }
        else
        {
            _editableResultsTable?.RejectChanges();
            SetStatus("Discarded unsaved row changes.");
        }

        await LoadEditableRowsAsync();
    }

    private bool HasPendingEditRowsChanges()
    {
        if (_editableResultsTable is null)
        {
            return false;
        }

        return _editableResultsTable.Rows
            .Cast<DataRow>()
            .Any(row => row.RowState is DataRowState.Added or DataRowState.Modified or DataRowState.Deleted);
    }

    private async Task LoadEditableRowsAsync()
    {
        if (!TryGetConnectionString(out var connectionString))
        {
            return;
        }

        if (_selectedTable is null)
        {
            SetStatus("Select a table first to enter edit mode.");
            return;
        }

        if (_selectedColumns.Count == 0)
        {
            SetStatus("Table metadata is not loaded yet. Select the table again and retry.");
            return;
        }

        var editQueryMode = QueryOutputModeParser.Parse(EditQueryTextBox.Text ?? string.Empty);
        _currentFullOutputMode = editQueryMode.HasFullDirective || FullOutputCheckBox.IsChecked == true;

        SyncInputsFromEditQueryText();

        var topRows = ParseEditTopRowsInput();
        var filter = string.IsNullOrWhiteSpace(EditFilterTextBox.Text) ? null : EditFilterTextBox.Text.Trim();
        var orderBy = string.IsNullOrWhiteSpace(EditOrderByTextBox.Text) ? null : EditOrderByTextBox.Text.Trim();

        UpdateEditQueryTextFromInputs();

        if (topRows > 5000)
        {
            topRows = 5000;
            _isSyncingEditQuery = true;
            EditTopRowsTextBox.Text = "5000";
            _isSyncingEditQuery = false;
            UpdateEditQueryTextFromInputs();
        }

        SetExecutionState(true);
        SetStatus(_currentFullOutputMode
            ? $"Loading editable rows from {_selectedTable.FullName} (full output mode)..."
            : $"Loading editable rows from {_selectedTable.FullName}...");

        try
        {
            _editableResultsTable = await _rowEditService.LoadTopRowsAsync(
                connectionString,
                _selectedTable.SchemaName,
                _selectedTable.TableName,
                topRows,
                filter,
                orderBy,
                ParseTimeoutSeconds(),
                CancellationToken.None);

            _isEditMode = true;
            _currentDataTable = _editableResultsTable;
            EditRowsDataGrid.ItemsSource = _editableResultsTable.DefaultView;
            EditRowsSummaryTextBlock.Text = $"Edit Mode ({_editableResultsTable.Rows.Count} rows loaded)";
            OutputTabControl.SelectedIndex = OutputEditRowsTabIndex;
            ApplyEditModeState();

            SetStatus($"Edit mode ready for {_selectedTable.FullName}.");
        }
        catch (Exception ex)
        {
            SetStatus($"Failed to load editable rows: {ex.Message}");
        }
        finally
        {
            SetExecutionState(false);
        }
    }

    private void ApplyEditModeState()
    {
        EditRowsDataGrid.IsReadOnly = !_isEditMode;
        EditRowsDataGrid.CanUserAddRows = _isEditMode;
        EditRowsDataGrid.CanUserDeleteRows = false;

        EnterEditModeButton.Content = _isEditMode ? "Reload Edit Rows" : "Edit Top Rows";
        SaveRowChangesButton.IsEnabled = _isEditMode;
        DiscardRowChangesButton.IsEnabled = _isEditMode;
        DeleteRowMenuItem.IsEnabled = _isEditMode;
        EditRowsSummaryTextBlock.Text = _isEditMode
            ? EditRowsSummaryTextBlock.Text
            : "Edit mode is not active.";
    }

    private void EditRowsInput_TextChanged(object sender, TextChangedEventArgs e)
    {
        if (_isSyncingEditQuery)
        {
            return;
        }

        UpdateEditQueryTextFromInputs();
    }

    private void EditQueryTextBox_TextChanged(object sender, TextChangedEventArgs e)
    {
        if (_isSyncingEditQuery)
        {
            return;
        }

        SyncInputsFromEditQueryText();
    }

    private void UpdateEditQueryTextFromInputs()
    {
        if (_isSyncingEditQuery)
        {
            return;
        }

        if (EditTopRowsTextBox is null || EditFilterTextBox is null || EditOrderByTextBox is null || EditQueryTextBox is null)
        {
            return;
        }

        var topRows = ParseEditTopRowsInput();
        var filter = string.IsNullOrWhiteSpace(EditFilterTextBox.Text) ? null : EditFilterTextBox.Text.Trim();
        var orderBy = string.IsNullOrWhiteSpace(EditOrderByTextBox.Text) ? null : EditOrderByTextBox.Text.Trim();

        var generatedSql = BuildEditRowsQuery(topRows, filter, orderBy);

        _isSyncingEditQuery = true;
        EditQueryTextBox.Text = generatedSql;
        EditQueryTextBox.CaretIndex = EditQueryTextBox.Text.Length;
        _isSyncingEditQuery = false;
    }

    private void SyncInputsFromEditQueryText()
    {
        if (_isSyncingEditQuery)
        {
            return;
        }

        if (EditTopRowsTextBox is null || EditFilterTextBox is null || EditOrderByTextBox is null || EditQueryTextBox is null)
        {
            return;
        }

        var sql = EditQueryTextBox.Text;
        if (!TryParseEditRowsQuery(sql, out var topRows, out var filter, out var orderBy))
        {
            return;
        }

        _isSyncingEditQuery = true;
        EditTopRowsTextBox.Text = topRows.ToString();
        EditFilterTextBox.Text = filter ?? string.Empty;
        EditOrderByTextBox.Text = orderBy ?? string.Empty;
        _isSyncingEditQuery = false;
    }

    private int ParseEditTopRowsInput()
    {
        if (EditTopRowsTextBox is null)
        {
            return 200;
        }

        if (int.TryParse(EditTopRowsTextBox.Text.Trim(), out var topRows) && topRows > 0)
        {
            return topRows;
        }

        _isSyncingEditQuery = true;
        EditTopRowsTextBox.Text = "200";
        _isSyncingEditQuery = false;
        return 200;
    }

    private string BuildEditRowsQuery(int topRows, string? filter, string? orderBy)
    {
        if (_selectedTable is null)
        {
            return "-- Select a table to generate an editable-row query.";
        }

        var sql = $"SELECT TOP ({topRows}) *{Environment.NewLine}FROM [{_selectedTable.SchemaName}].[{_selectedTable.TableName}]";

        if (!string.IsNullOrWhiteSpace(filter))
        {
            sql += $"{Environment.NewLine}WHERE {filter}";
        }

        if (!string.IsNullOrWhiteSpace(orderBy))
        {
            sql += $"{Environment.NewLine}ORDER BY {orderBy}";
        }

        sql += ";";
        return sql;
    }

    private static bool TryParseEditRowsQuery(string sql, out int topRows, out string? filter, out string? orderBy)
    {
        topRows = 200;
        filter = null;
        orderBy = null;

        var match = EditRowsQueryRegex.Match(sql ?? string.Empty);
        if (!match.Success)
        {
            return false;
        }

        if (!int.TryParse(match.Groups["top"].Value, out topRows) || topRows <= 0)
        {
            return false;
        }

        filter = match.Groups["where"].Success
            ? match.Groups["where"].Value.Trim()
            : null;

        orderBy = match.Groups["order"].Success
            ? match.Groups["order"].Value.Trim().TrimEnd(';')
            : null;

        if (string.IsNullOrWhiteSpace(filter))
        {
            filter = null;
        }

        if (string.IsNullOrWhiteSpace(orderBy))
        {
            orderBy = null;
        }

        return true;
    }

    private static IReadOnlyList<RowUpdateRequest> BuildRowUpdates(DataTable table, IReadOnlyList<ColumnSchemaInfo> columns)
    {
        var columnNames = columns.Select(c => c.ColumnName).ToHashSet(StringComparer.OrdinalIgnoreCase);
        var primaryKeyNames = columns.Where(c => c.IsPrimaryKey).Select(c => c.ColumnName).ToList();

        var updates = new List<RowUpdateRequest>();
        foreach (DataRow row in table.Rows)
        {
            if (row.RowState != DataRowState.Modified)
            {
                continue;
            }

            var keyValues = new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase);
            foreach (var keyName in primaryKeyNames)
            {
                if (!table.Columns.Contains(keyName))
                {
                    continue;
                }

                var value = row[keyName, DataRowVersion.Original];
                keyValues[keyName] = value == DBNull.Value ? null : value;
            }

            var currentValues = new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase);
            foreach (DataColumn column in table.Columns)
            {
                if (!columnNames.Contains(column.ColumnName))
                {
                    continue;
                }

                var value = row[column, DataRowVersion.Current];
                currentValues[column.ColumnName] = value == DBNull.Value ? null : value;
            }

            updates.Add(new RowUpdateRequest
            {
                OriginalKeyValues = keyValues,
                CurrentValues = currentValues
            });
        }

        return updates;
    }

    private static IReadOnlyList<RowInsertRequest> BuildRowInserts(DataTable table, IReadOnlyList<ColumnSchemaInfo> columns)
    {
        var insertableColumns = columns
            .Where(IsInsertableColumn)
            .Select(c => c.ColumnName)
            .ToHashSet(StringComparer.OrdinalIgnoreCase);

        var inserts = new List<RowInsertRequest>();
        foreach (DataRow row in table.Rows)
        {
            if (row.RowState != DataRowState.Added)
            {
                continue;
            }

            var values = new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase);
            foreach (DataColumn column in table.Columns)
            {
                if (!insertableColumns.Contains(column.ColumnName))
                {
                    continue;
                }

                var value = row[column, DataRowVersion.Current];
                values[column.ColumnName] = value == DBNull.Value ? null : value;
            }

            if (values.Count == 0)
            {
                continue;
            }

            inserts.Add(new RowInsertRequest
            {
                Values = values
            });
        }

        return inserts;
    }

    private static bool IsInsertableColumn(ColumnSchemaInfo column)
    {
        if (column.IsIdentity)
        {
            return false;
        }

        return !column.DataType.Equals("rowversion", StringComparison.OrdinalIgnoreCase)
            && !column.DataType.Equals("timestamp", StringComparison.OrdinalIgnoreCase);
    }

    private static IReadOnlyDictionary<string, object?> BuildPrimaryKeyValues(DataRow row, IReadOnlyList<ColumnSchemaInfo> columns)
    {
        var keyValues = new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase);
        foreach (var keyColumn in columns.Where(c => c.IsPrimaryKey))
        {
            if (!row.Table.Columns.Contains(keyColumn.ColumnName))
            {
                continue;
            }

            object? value;
            if (row.RowState == DataRowState.Modified)
            {
                value = row[keyColumn.ColumnName, DataRowVersion.Original];
            }
            else
            {
                value = row[keyColumn.ColumnName];
            }

            keyValues[keyColumn.ColumnName] = value == DBNull.Value ? null : value;
        }

        return keyValues;
    }

    private bool TryPromptForQueryParameters(
        IReadOnlyList<string> parameterNames,
        out IReadOnlyList<QueryParameterValue> parameters)
    {
        var rows = new ObservableCollection<QueryParameterEditorRow>(
            parameterNames.Select(name => new QueryParameterEditorRow
            {
                ParameterName = name,
                Value = string.Empty,
                SendAsNull = false
            }));

        var dialog = new Window
        {
            Title = "Query Parameters",
            Owner = this,
            WindowStartupLocation = WindowStartupLocation.CenterOwner,
            Width = 620,
            Height = 420,
            MinWidth = 560,
            MinHeight = 340,
            ResizeMode = ResizeMode.CanResize,
            ShowInTaskbar = false,
            Background = (Brush)FindResource("SurfaceBrush"),
            Foreground = (Brush)FindResource("TextPrimaryBrush")
        };

        var root = new Grid
        {
            Margin = new Thickness(14)
        };
        root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
        root.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
        root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

        var instructions = new TextBlock
        {
            Text = "Enter values for SQL variables before execution.",
            Margin = new Thickness(0, 0, 0, 8),
            FontWeight = FontWeights.SemiBold
        };
        Grid.SetRow(instructions, 0);
        root.Children.Add(instructions);

        var dataGrid = new DataGrid
        {
            AutoGenerateColumns = false,
            CanUserAddRows = false,
            CanUserDeleteRows = false,
            HeadersVisibility = DataGridHeadersVisibility.All,
            ItemsSource = rows
        };
        dataGrid.Columns.Add(new DataGridTextColumn
        {
            Header = "Variable",
            Binding = new Binding(nameof(QueryParameterEditorRow.ParameterName)),
            IsReadOnly = true,
            Width = new DataGridLength(180)
        });
        dataGrid.Columns.Add(new DataGridTextColumn
        {
            Header = "Value",
            Binding = new Binding(nameof(QueryParameterEditorRow.Value)) { UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged },
            Width = new DataGridLength(1, DataGridLengthUnitType.Star)
        });
        dataGrid.Columns.Add(new DataGridCheckBoxColumn
        {
            Header = "NULL",
            Binding = new Binding(nameof(QueryParameterEditorRow.SendAsNull)),
            Width = new DataGridLength(80)
        });
        Grid.SetRow(dataGrid, 1);
        root.Children.Add(dataGrid);

        var buttons = new StackPanel
        {
            Orientation = Orientation.Horizontal,
            HorizontalAlignment = HorizontalAlignment.Right,
            Margin = new Thickness(0, 10, 0, 0)
        };
        var cancelButton = new Button
        {
            Content = "Cancel",
            MinWidth = 90,
            Margin = new Thickness(0, 0, 8, 0),
            IsCancel = true
        };
        var executeButton = new Button
        {
            Content = "Execute",
            MinWidth = 90,
            IsDefault = true
        };
        executeButton.Click += (_, _) => dialog.DialogResult = true;
        buttons.Children.Add(cancelButton);
        buttons.Children.Add(executeButton);
        Grid.SetRow(buttons, 2);
        root.Children.Add(buttons);

        dialog.Content = root;

        if (dialog.ShowDialog() != true)
        {
            parameters = Array.Empty<QueryParameterValue>();
            return false;
        }

        parameters = rows
            .Select(row => new QueryParameterValue
            {
                Name = row.ParameterName,
                Value = row.SendAsNull ? null : row.Value
            })
            .ToList();

        return true;
    }

    private sealed class ProcedureParameterEditorRow
    {
        public required string ParameterName { get; init; }

        public required string DataType { get; init; }

        public string? Value { get; set; }

        public bool SendAsNull { get; set; }

        public bool IsOutput { get; init; }
    }

    private sealed class QueryParameterEditorRow
    {
        public required string ParameterName { get; init; }

        public string? Value { get; set; }

        public bool SendAsNull { get; set; }
    }
}