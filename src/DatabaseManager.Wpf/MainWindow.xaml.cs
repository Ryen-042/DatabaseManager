using System.Data;
using System.Collections.ObjectModel;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Interop;
using System.Windows.Input;
using System.Windows.Media;
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

    private readonly IDatabaseQueryService _databaseQueryService = new SqlServerQueryService();
    private readonly ITemplateStoreService _templateStoreService = new TemplateStoreService();
    private readonly IExportService _exportService = new ExportService();
    private readonly IDatabaseSchemaService _databaseSchemaService = new SqlServerSchemaService();
    private readonly IQueryAssistantService _queryAssistantService = new SqlQueryAssistantService();
    private readonly IStoredProcedureExecutionService _storedProcedureExecutionService = new StoredProcedureExecutionService();

    private DataTable? _currentDataTable;
    private CancellationTokenSource? _executionCancellationTokenSource;
    private List<TableSchemaInfo> _tables = new();
    private List<StoredProcedureSchemaInfo> _storedProcedures = new();
    private List<ColumnSchemaInfo> _selectedColumns = new();
    private List<StoredProcedureParameterInfo> _selectedProcedureParameters = new();
    private TableSchemaInfo? _selectedTable;
    private StoredProcedureSchemaInfo? _selectedStoredProcedure;
    private readonly ObservableCollection<ProcedureParameterEditorRow> _runnerParameterRows = new();

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

    private async void TestConnectionButton_Click(object sender, RoutedEventArgs e)
    {
        if (!TryGetConnectionString(out var connectionString))
        {
            return;
        }

        SetExecutionState(true);
        SetStatus("Testing database connection...");

        using var cts = new CancellationTokenSource();
        var result = await _databaseQueryService.ExecuteAsync(
            connectionString,
            "SELECT 1 AS IsConnected;",
            ParseTimeoutSeconds(),
            cts.Token);

        if (result.IsSuccess)
        {
            SetStatus("Connection test succeeded.");
            await LoadSchemaMetadataAsync(connectionString);
        }
        else
        {
            SetStatus($"Connection test failed: {result.ErrorMessage}");
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

        _executionCancellationTokenSource?.Dispose();
        _executionCancellationTokenSource = new CancellationTokenSource();

        SetExecutionState(true);
        SetStatus("Executing query...");

        var result = await _databaseQueryService.ExecuteAsync(
            connectionString,
            sql,
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

        await _exportService.ExportToCsvAsync(_currentDataTable!, dialog.FileName, CancellationToken.None);
        SetStatus($"CSV export complete: {dialog.FileName}");
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

        await _exportService.ExportToExcelAsync(_currentDataTable!, dialog.FileName, CancellationToken.None);
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
        SetStatus("Generated SELECT query from table metadata.");
    }

    private void GenerateInsertButton_Click(object sender, RoutedEventArgs e)
    {
        if (!EnsureTableSelection())
        {
            return;
        }

        QueryTextBox.Text = _queryAssistantService.BuildInsertQuery(_selectedTable!, _selectedColumns);
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
        SetStatus("Generated table SQL schema text.");
    }

    private void GenerateExecButton_Click(object sender, RoutedEventArgs e)
    {
        if (!EnsureProcedureSelection())
        {
            return;
        }

        QueryTextBox.Text = _queryAssistantService.BuildExecuteProcedureQuery(_selectedStoredProcedure!, _selectedProcedureParameters);
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
        SchemaAssistantTabControl.SelectedIndex = 2;
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

        SetExecutionState(true);
        SetStatus("Executing stored procedure...");

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
        if (!result.IsSuccess)
        {
            SetStatus($"{operationName} failed after {result.Duration.TotalSeconds:F2}s: {result.ErrorMessage}");
            return;
        }

        _currentDataTable = result.DataTable;
        ResultsDataGrid.ItemsSource = _currentDataTable?.DefaultView;
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
        TestConnectionButton.IsEnabled = !isExecuting;
        CancelButton.IsEnabled = isExecuting;
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

        var copiedText = TryGetClickedCellText(clickedCell);
        if (string.IsNullOrWhiteSpace(copiedText))
        {
            return;
        }

        Clipboard.SetText(copiedText);
        SetStatus($"Copied value to clipboard: {TruncateForStatus(copiedText)}");
    }

    private void SchemaListBox_PreviewMouseLeftButtonUp(object sender, MouseButtonEventArgs e)
    {
        if (sender is not ListBox listBox)
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

        Clipboard.SetText(copiedText);
        SetStatus($"Copied value to clipboard: {TruncateForStatus(copiedText)}");
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
            return value is null or DBNull ? string.Empty : value.ToString();
        }

        var property = item.GetType().GetProperty(propertyPath);
        if (property is null)
        {
            return null;
        }

        var propertyValue = property.GetValue(item);
        return propertyValue?.ToString();
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

    private sealed class ProcedureParameterEditorRow
    {
        public required string ParameterName { get; init; }

        public required string DataType { get; init; }

        public string? Value { get; set; }

        public bool SendAsNull { get; set; }

        public bool IsOutput { get; init; }
    }
}