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
using System.Windows.Controls.Primitives;
using System.Windows.Threading;
using DatabaseManager.Core.Models.Editing;
using DatabaseManager.Core.Models;
using DatabaseManager.Core.Models.Schema;
using DatabaseManager.Core.Services;
using DatabaseManager.Core.Services.Schema;
using DatabaseManager.Wpf.Editors;
using DatabaseManager.Wpf.SqlSuggestions;
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
    private const int MaxRecentSqlFragments = 20;

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
    private List<ForeignKeySchemaInfo> _foreignKeys = new();
    private List<ColumnSchemaInfo> _selectedColumns = new();
    private List<StoredProcedureParameterInfo> _selectedProcedureParameters = new();
    private TableSchemaInfo? _selectedTable;
    private StoredProcedureSchemaInfo? _selectedStoredProcedure;
    private readonly ObservableCollection<ProcedureParameterEditorRow> _runnerParameterRows = new();
    private readonly GridLength _schemaPaneExpandedWidth = new(330);
    private DataTable? _editableResultsTable;
    private bool _isEditMode;
    private bool _isSyncingEditQuery;
    private bool _isEditRowsCustomQueryMode;
    private int _lastEditRowsCurrentRowIndex = -1;
    private readonly ISqlSuggestionEngine _sqlSuggestionEngine = new SqlSuggestionEngine();
    private readonly ISqlCompletionCatalogService _sqlCompletionCatalogService = new SqlCompletionCatalogService();
    private ISqlTextEditor? _sqlEditor;
    private ISqlTextEditor? _editRowsSqlEditor;
    private CancellationTokenSource? _sqlSuggestionDebounceCts;
    private ISqlTextEditor? _activeSqlSuggestionTextEditor;
    private int _activeSqlSuggestionTokenStart;
    private int _activeSqlSuggestionTokenLength;
    private readonly LinkedList<string> _recentSqlFragments = new();
    private readonly HashSet<string> _recentSqlFragmentLookup = new(StringComparer.OrdinalIgnoreCase);

    private static readonly Regex EditRowsQueryRegex = new(
        @"^\s*SELECT\s+(?:TOP\s*\(\s*(?<top>\d+)\s*\)\s+)?\*\s+FROM\s+(?<from>\[[^\]]+\]\.\[[^\]]+\]|\S+)(?:\s+WHERE\s+(?<where>.*?))?(?:\s+ORDER\s+BY\s+(?<order>.*?))?\s*;?\s*$",
        RegexOptions.IgnoreCase | RegexOptions.Singleline | RegexOptions.Compiled);

    public MainWindow()
    {
        InitializeComponent();
        _sqlEditor = new AvalonEditSqlTextEditorAdapter(QueryTextBox);
        _editRowsSqlEditor = new AvalonEditSqlTextEditorAdapter(EditQueryTextBox);
        SqlEditorSupport.Configure(QueryTextBox);
        SqlEditorSupport.Configure(EditQueryTextBox);
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
        ApplyEditRowsCornerButtonStyle();
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

        if (TryGetOutputTabShortcutIndex(e.Key, out var tabIndex))
        {
            OutputTabControl.SelectedIndex = tabIndex;
            e.Handled = true;
            return;
        }

        if (OutputTabControl.SelectedIndex == OutputEditRowsTabIndex)
        {
            if (e.Key == Key.R)
            {
                await RefreshEditRowsAsync();
                e.Handled = true;
                return;
            }

            if (e.Key == Key.S)
            {
                await SaveRowChangesAsync();
                e.Handled = true;
                return;
            }
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

    private static bool TryGetOutputTabShortcutIndex(Key key, out int tabIndex)
    {
        tabIndex = key switch
        {
            Key.D1 or Key.NumPad1 => 0,
            Key.D2 or Key.NumPad2 => 1,
            Key.D3 or Key.NumPad3 => 2,
            Key.D4 or Key.NumPad4 => 3,
            Key.D5 or Key.NumPad5 => 4,
            _ => -1
        };

        return tabIndex >= 0;
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
        ApplyEditRowsCornerButtonStyle();
    }

    private void DarkModeCheckBox_Unchecked(object sender, RoutedEventArgs e)
    {
        App.ApplyTheme(false);
        ApplyTitleBarTheme(false);
        ApplyEditRowsCornerButtonStyle();
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
        if (!TryGetConnectionString(out var connectionString, showMissingStatus: !triggeredOnStartup))
        {
            if (triggeredOnStartup)
            {
                SetStatus("Startup auto-connect skipped: connection string is empty.");
            }

            return;
        }

        SetExecutionState(true);
        SetStatus(triggeredOnStartup ? "Attempting startup database connection..." : "Connecting to database...");

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
        TrackRecentSqlFragments(sqlToExecute);

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
        TrackRecentSqlFragments(selected.Sql);
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
        await RefreshEditRowsAsync();
    }

    private async void ApplyEditFilterButton_Click(object sender, RoutedEventArgs e)
    {
        await RefreshEditRowsAsync();
    }

    private void EditRowsDataGrid_PreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
    {
        if (!_isEditMode)
        {
            return;
        }

        var checkBox = FindAncestor<CheckBox>(e.OriginalSource as DependencyObject);
        if (checkBox is null)
        {
            return;
        }

        var cell = FindAncestor<DataGridCell>(checkBox);
        var row = FindAncestor<DataGridRow>(checkBox);
        if (cell is null || row is null || cell.IsReadOnly)
        {
            return;
        }

        EditRowsDataGrid.SelectedItem = row.Item;
        EditRowsDataGrid.CurrentCell = new DataGridCellInfo(row.Item, cell.Column);
        if (!cell.IsEditing)
        {
            EditRowsDataGrid.BeginEdit(e);
        }

        checkBox.IsChecked = !(checkBox.IsChecked ?? false);

        var bindingExpression = BindingOperations.GetBindingExpression(checkBox, ToggleButton.IsCheckedProperty);
        bindingExpression?.UpdateSource();

        EditRowsDataGrid.CommitEdit(DataGridEditingUnit.Cell, true);
        EditRowsDataGrid.CommitEdit(DataGridEditingUnit.Row, true);
        RefreshEditRowsVisualStates();

        e.Handled = true;
    }

    private void EditRowsDataGrid_PreviewKeyDown(object sender, KeyEventArgs e)
    {
        if (e.Key != Key.Delete)
        {
            return;
        }

        if (!_isEditMode || _editableResultsTable is null)
        {
            return;
        }

        DeleteRowMenuItem_Click(this, new RoutedEventArgs());
        e.Handled = true;
    }

    private void EditRowsDataGrid_LoadingRow(object sender, DataGridRowEventArgs e)
    {
        ApplyEditRowsRowVisualState(e.Row);
        ApplyEditRowsCornerButtonStyle();
    }

    private void DataGrid_Loaded(object sender, RoutedEventArgs e)
    {
        if (sender is DataGrid dataGrid)
        {
            Dispatcher.BeginInvoke(DispatcherPriority.Background, new Action(() =>
            {
                ResetDataGridHorizontalScroll(dataGrid);
            }));
        }
    }

    private void ResetDataGridHorizontalScroll(DataGrid dataGrid)
    {
        var scrollViewer = FindDescendant<ScrollViewer>(dataGrid);
        if (scrollViewer != null)
        {
            scrollViewer.ScrollToHorizontalOffset(0);
        }
    }

    private void EditRowsDataGrid_CurrentCellChanged(object? sender, EventArgs e)
    {
        if (EditRowsDataGrid.Items.Count == 0)
        {
            _lastEditRowsCurrentRowIndex = -1;
            return;
        }

        var currentIndex = EditRowsDataGrid.CurrentItem is null
            ? -1
            : EditRowsDataGrid.Items.IndexOf(EditRowsDataGrid.CurrentItem);

        ApplyEditRowsVisualStateAtIndex(_lastEditRowsCurrentRowIndex);
        ApplyEditRowsVisualStateAtIndex(currentIndex);

        _lastEditRowsCurrentRowIndex = currentIndex;
    }

    private void EditRowsDataGrid_CellEditEnding(object? sender, DataGridCellEditEndingEventArgs e)
    {
        Dispatcher.BeginInvoke(DispatcherPriority.Background, new Action(RefreshEditRowsVisualStates));
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
            RefreshEditRowsVisualStates();
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
        RefreshEditRowsVisualStates();
    }

    private void ResultsDataGrid_PreviewMouseRightButtonDown(object sender, MouseButtonEventArgs e)
    {
        if (sender is not DataGrid dataGrid)
        {
            return;
        }

        var cell = FindAncestor<DataGridCell>(e.OriginalSource as DependencyObject);
        if (cell is not null)
        {
            var rowItem = cell.DataContext;
            if (rowItem is not null)
            {
                dataGrid.CurrentCell = new DataGridCellInfo(rowItem, cell.Column);
                dataGrid.SelectedItem = rowItem;
                cell.Focus();
                return;
            }
        }

        var row = FindAncestor<DataGridRow>(e.OriginalSource as DependencyObject);
        if (row is null)
        {
            return;
        }

        row.IsSelected = true;
        dataGrid.SelectedItem = row.Item;
    }

    private async void CopyDataGridCellValueMenuItem_Click(object sender, RoutedEventArgs e)
    {
        if (sender is not MenuItem menuItem
            || menuItem.Parent is not ContextMenu contextMenu
            || contextMenu.PlacementTarget is not DataGrid dataGrid)
        {
            return;
        }

        if (!TryGetCurrentDataGridCellTextForCopy(dataGrid, out var copiedText)
            || string.IsNullOrWhiteSpace(copiedText))
        {
            SetStatus("Select a cell value to copy.");
            return;
        }

        if (await TrySetClipboardTextAsync(copiedText))
        {
            SetStatus($"Copied value to clipboard: {TruncateForStatus(copiedText)}");
        }
        else
        {
            SetStatus("Could not access the clipboard. Please try again.");
        }
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

        var selectedRowViews = GetSelectedEditRowsSelection();
        if (selectedRowViews.Count == 0)
        {
            SetStatus("Select a row to delete.");
            return;
        }

        if (_selectedColumns.All(c => !c.IsPrimaryKey))
        {
            await DeleteRowsWithoutPrimaryKeyAsync(connectionString, selectedRowViews);
            return;
        }

        await DeleteRowsWithPrimaryKeyAsync(connectionString, selectedRowViews);
    }

    private async Task DeleteRowsWithPrimaryKeyAsync(string connectionString, IReadOnlyList<DataRowView> selectedRowViews)
    {
        if (_selectedTable is null || _editableResultsTable is null)
        {
            return;
        }

        var firstKeyValues = BuildPrimaryKeyValues(selectedRowViews[0].Row, _selectedColumns);
        var keyDetails = string.Join(", ", firstKeyValues.Select(x => $"{x.Key}={x.Value}"));

        var confirmation = MessageBox.Show(
            selectedRowViews.Count == 1
                ? $"Delete this row from {_selectedTable.FullName}?{Environment.NewLine}{Environment.NewLine}{keyDetails}"
                : $"Delete {selectedRowViews.Count} selected rows from {_selectedTable.FullName}?",
            "Delete Row",
            MessageBoxButton.YesNo,
            MessageBoxImage.Warning);

        if (confirmation != MessageBoxResult.Yes)
        {
            return;
        }

        SetExecutionState(true);
        SetStatus("Deleting selected row(s)...");

        try
        {
            var removedRows = new List<DataRow>();

            foreach (var rowView in selectedRowViews)
            {
                var keyValues = BuildPrimaryKeyValues(rowView.Row, _selectedColumns);
                var affectedRows = await _rowEditService.DeleteRowAsync(
                    connectionString,
                    _selectedTable.SchemaName,
                    _selectedTable.TableName,
                    _selectedColumns,
                    keyValues,
                    ParseTimeoutSeconds(),
                    CancellationToken.None);

                if (affectedRows > 0)
                {
                    removedRows.Add(rowView.Row);
                }
            }

            foreach (var row in removedRows)
            {
                _editableResultsTable.Rows.Remove(row);
            }

            _editableResultsTable.AcceptChanges();
            SetStatus($"Deleted {removedRows.Count} row(s) successfully.");
            RefreshEditRowsVisualStates();
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

    private async Task DeleteRowsWithoutPrimaryKeyAsync(string connectionString, IReadOnlyList<DataRowView> selectedRowViews)
    {
        if (_selectedTable is null || _editableResultsTable is null)
        {
            return;
        }

        if (!TryPromptForDeleteColumns(_selectedColumns, out var selectedColumns))
        {
            SetStatus("Delete canceled.");
            return;
        }

        if (selectedColumns.Count == 0)
        {
            SetStatus("Select at least one column to build delete filters.");
            return;
        }

        var confirmation = MessageBox.Show(
            $"Delete {selectedRowViews.Count} selected row(s) from {_selectedTable.FullName} using filters on: {string.Join(", ", selectedColumns)}?",
            "Delete Rows Without Primary Key",
            MessageBoxButton.YesNo,
            MessageBoxImage.Warning);

        if (confirmation != MessageBoxResult.Yes)
        {
            return;
        }

        var rowValues = selectedRowViews
            .Select(rowView => BuildSelectedColumnValues(rowView.Row, selectedColumns))
            .ToList();

        SetExecutionState(true);
        SetStatus("Validating delete impact...");

        try
        {
            var result = await _rowEditService.DeleteRowsBySelectedColumnsAsync(
                connectionString,
                _selectedTable.SchemaName,
                _selectedTable.TableName,
                selectedColumns,
                rowValues,
                ParseTimeoutSeconds(),
                CancellationToken.None);

            if (!result.WasExecuted)
            {
                var warning = MessageBox.Show(
                    $"Delete validation mismatch: intended {result.IntendedRows} row(s), matched {result.MatchedRows}.{Environment.NewLine}{Environment.NewLine}Copy generated DELETE SQL for manual review?",
                    "Delete Validation Warning",
                    MessageBoxButton.YesNo,
                    MessageBoxImage.Warning);

                if (warning == MessageBoxResult.Yes)
                {
                    if (await TrySetClipboardTextAsync(result.GeneratedDeleteSql))
                    {
                        SetStatus("Delete validation mismatch. Generated SQL copied to clipboard.");
                    }
                    else
                    {
                        SetStatus("Delete validation mismatch and clipboard was unavailable.");
                    }
                }
                else
                {
                    SetStatus($"Delete canceled due to mismatch (intended {result.IntendedRows}, matched {result.MatchedRows}).");
                }

                return;
            }

            foreach (var rowView in selectedRowViews)
            {
                _editableResultsTable.Rows.Remove(rowView.Row);
            }

            _editableResultsTable.AcceptChanges();
            SetStatus($"Deleted {result.DeletedRows} row(s) successfully.");
            RefreshEditRowsVisualStates();
        }
        catch (Exception ex)
        {
            SetStatus($"Failed to delete rows: {ex.Message}");
        }
        finally
        {
            SetExecutionState(false);
        }
    }

    private List<DataRowView> GetSelectedEditRowsSelection()
    {
        var rows = EditRowsDataGrid.SelectedItems
            .OfType<DataRowView>()
            .Distinct()
            .ToList();

        if (rows.Count == 0 && EditRowsDataGrid.SelectedItem is DataRowView singleRow)
        {
            rows.Add(singleRow);
        }

        return rows;
    }

    private static IReadOnlyDictionary<string, object?> BuildSelectedColumnValues(DataRow row, IReadOnlyList<string> selectedColumns)
    {
        var values = new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase);
        foreach (var column in selectedColumns)
        {
            if (!row.Table.Columns.Contains(column))
            {
                throw new InvalidOperationException($"Selected row does not contain column '{column}'.");
            }

            var value = row.RowState == DataRowState.Modified
                ? row[column, DataRowVersion.Original]
                : row[column];

            values[column] = value == DBNull.Value ? null : value;
        }

        return values;
    }

    private bool TryPromptForDeleteColumns(
        IReadOnlyList<ColumnSchemaInfo> availableColumns,
        out IReadOnlyList<string> selectedColumns)
    {
        var rows = new ObservableCollection<DeleteColumnSelectionRow>(
            availableColumns
                .OrderBy(c => c.OrdinalPosition)
                .Select(c => new DeleteColumnSelectionRow
                {
                    ColumnName = c.ColumnName,
                    DataType = c.DataType,
                    IsSelected = c.IsPrimaryKey
                }));

        var dialog = new Window
        {
            Title = "Select Columns For Delete Filter",
            Owner = this,
            WindowStartupLocation = WindowStartupLocation.CenterOwner,
            Width = 620,
            Height = 480,
            MinWidth = 560,
            MinHeight = 360,
            ResizeMode = ResizeMode.CanResize,
            ShowInTaskbar = false,
            Background = (Brush)FindResource("SurfaceBrush"),
            Foreground = (Brush)FindResource("TextPrimaryBrush")
        };

        var root = new Grid { Margin = new Thickness(14) };
        root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
        root.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
        root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

        var instructions = new TextBlock
        {
            Text = "Select columns to build WHERE predicates for each selected row.",
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
        dataGrid.Columns.Add(new DataGridCheckBoxColumn
        {
            Header = "Use",
            Binding = new Binding(nameof(DeleteColumnSelectionRow.IsSelected)),
            Width = new DataGridLength(70)
        });
        dataGrid.Columns.Add(new DataGridTextColumn
        {
            Header = "Column",
            Binding = new Binding(nameof(DeleteColumnSelectionRow.ColumnName)),
            IsReadOnly = true,
            Width = new DataGridLength(1, DataGridLengthUnitType.Star)
        });
        dataGrid.Columns.Add(new DataGridTextColumn
        {
            Header = "Data Type",
            Binding = new Binding(nameof(DeleteColumnSelectionRow.DataType)),
            IsReadOnly = true,
            Width = new DataGridLength(150)
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
        var applyButton = new Button
        {
            Content = "Apply",
            MinWidth = 90,
            IsDefault = true
        };
        applyButton.Click += (_, _) => dialog.DialogResult = true;

        buttons.Children.Add(cancelButton);
        buttons.Children.Add(applyButton);
        Grid.SetRow(buttons, 2);
        root.Children.Add(buttons);

        dialog.Content = root;

        if (dialog.ShowDialog() != true)
        {
            selectedColumns = Array.Empty<string>();
            return false;
        }

        selectedColumns = rows
            .Where(row => row.IsSelected)
            .Select(row => row.ColumnName)
            .ToList();

        return true;
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
            ExitEditModeAndClearEditableRows();
            EditRowsDataGrid.ItemsSource = null;
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
            _sqlCompletionCatalogService.RefreshTableColumns(_selectedTable, _selectedColumns);
            TableColumnsDataGrid.ItemsSource = _selectedColumns;
            TableSqlDefinitionTextBox.Text = _queryAssistantService.BuildTableSchemaText(_selectedTable, _selectedColumns);
            SelectedTableTextBlock.Text = $"{_selectedTable.FullName} ({_selectedColumns.Count} columns)";
            SchemaSummaryTextBlock.Text = $"Table selected: {_selectedTable.FullName}";
            UpdateSchemaDetailsPanelByAssistantTab();
            UpdateEditQueryTextFromInputs();
            ShowSelectedTableColumnsInEditRowsGrid();
            OutputTabControl.SelectedIndex = OutputEditRowsTabIndex;

            // Auto-substitute TableName placeholder with selected table name
            SubstituteTableNamePlaceholder(_selectedTable);
        }
        catch (Exception ex)
        {
            SetStatus($"Failed to load columns: {ex.Message}");
        }
    }

    private void SubstituteTableNamePlaceholder(TableSchemaInfo selectedTable)
    {
        var editors = new[] { _sqlEditor, _editRowsSqlEditor };
        foreach (var editor in editors)
        {
            if (editor?.Text?.Contains("TableName", StringComparison.Ordinal) != true)
            {
                continue;
            }

            var newText = Regex.Replace(
                editor.Text,
                @"\[TableName\]|TableName",
                $"[{selectedTable.SchemaName}].[{selectedTable.TableName}]",
                RegexOptions.Compiled);

            if (!string.Equals(editor.Text, newText, StringComparison.Ordinal))
            {
                var caretPos = editor.CaretIndex;
                editor.Text = newText;
                editor.CaretIndex = Math.Min(caretPos, editor.Text.Length);
            }
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
            _sqlCompletionCatalogService.RefreshProcedureParameters(_selectedStoredProcedure, _selectedProcedureParameters);
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

    private void GenerateDropTableScriptButton_Click(object sender, RoutedEventArgs e)
    {
        if (!EnsureTableSelection())
        {
            return;
        }

        QueryTextBox.Text = _queryAssistantService.BuildDropTableScript(_selectedTable!);
        OutputTabControl.SelectedIndex = OutputSqlEditorTabIndex;
        SetStatus("Generated DROP TABLE script in SQL Editor.");
    }

    private void GenerateDropAndCreateTableScriptButton_Click(object sender, RoutedEventArgs e)
    {
        if (!EnsureTableSelection())
        {
            return;
        }

        QueryTextBox.Text = _queryAssistantService.BuildDropAndCreateTableScript(_selectedTable!, _selectedColumns);
        OutputTabControl.SelectedIndex = OutputSqlEditorTabIndex;
        SetStatus("Generated DROP + CREATE TABLE script in SQL Editor.");
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

    private void GenerateDropProcedureScriptButton_Click(object sender, RoutedEventArgs e)
    {
        if (!EnsureProcedureSelection())
        {
            return;
        }

        QueryTextBox.Text = _queryAssistantService.BuildDropProcedureScript(_selectedStoredProcedure!);
        OutputTabControl.SelectedIndex = OutputSqlEditorTabIndex;
        SetStatus("Generated DROP PROCEDURE script in SQL Editor.");
    }

    private void GenerateAlterProcedureScriptButton_Click(object sender, RoutedEventArgs e)
    {
        if (!EnsureProcedureSelection())
        {
            return;
        }

        var definition = string.IsNullOrWhiteSpace(ProcedureSqlDefinitionTextBox.Text)
            ? null
            : ProcedureSqlDefinitionTextBox.Text;

        QueryTextBox.Text = _queryAssistantService.BuildAlterProcedureScript(_selectedStoredProcedure!, definition);
        OutputTabControl.SelectedIndex = OutputSqlEditorTabIndex;
        SetStatus("Generated ALTER PROCEDURE script in SQL Editor.");
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
            var foreignKeysTask = _databaseSchemaService.GetForeignKeysAsync(connectionString, CancellationToken.None);

            await Task.WhenAll(tablesTask, proceduresTask, foreignKeysTask);

            _tables = tablesTask.Result.ToList();
            _storedProcedures = proceduresTask.Result.ToList();
            _foreignKeys = foreignKeysTask.Result.ToList();
            _sqlCompletionCatalogService.RefreshSchemaMetadata(_tables, _storedProcedures, _foreignKeys);

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
            ExitEditModeAndClearEditableRows();
            EditRowsDataGrid.ItemsSource = null;

            ApplyTableFilter();
            ApplyProcedureFilter();

            if (_tables.Count == 0 && _storedProcedures.Count == 0)
            {
                SetStatus("Schema refresh completed, but no tables or procedures were found. Verify the target database in your connection string and user permissions.");
                return;
            }

            SetStatus($"Schema loaded successfully: {_tables.Count} tables, {_storedProcedures.Count} procedures, {_foreignKeys.Count} foreign keys.");
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

    private bool TryGetConnectionString(out string connectionString, bool showMissingStatus = true)
    {
        connectionString = ConnectionStringTextBox.Text.Trim();
        if (string.IsNullOrWhiteSpace(connectionString))
        {
            if (showMissingStatus)
            {
                SetStatus("Connection string is required.");
            }

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

        var resultSets = result.ResultSets?.Count > 0
            ? result.ResultSets
            : BuildFallbackResultSets(result);

        _currentDataTable = resultSets.FirstOrDefault(x => x.DataTable is not null)?.DataTable;
        RenderResultsSections(resultSets);
        OutputTabControl.SelectedIndex = OutputResultsTabIndex;
        ResultsSummaryTextBlock.Text = BuildResultsSummaryText(resultSets, result.AffectedRows);

        if (result.OutputParameters is { Count: > 0 })
        {
            var outputs = string.Join(", ", result.OutputParameters.Select(x => $"{x.Key}={x.Value ?? "NULL"}"));
            SetStatus($"{operationName} completed in {result.Duration.TotalSeconds:F2}s. Output: {outputs}");
            return;
        }

        SetStatus($"{operationName} completed in {result.Duration.TotalSeconds:F2}s.");
    }

    private IReadOnlyList<QueryResultSet> BuildFallbackResultSets(QueryExecutionResult result)
    {
        if (result.DataTable is not null)
        {
            return new[]
            {
                new QueryResultSet
                {
                    Title = "Result Set 1",
                    DataTable = result.DataTable,
                    AffectedRows = result.DataTable.Rows.Count
                }
            };
        }

        return new[]
        {
            new QueryResultSet
            {
                Title = "Statement Summary",
                DataTable = null,
                AffectedRows = result.AffectedRows
            }
        };
    }

    private void RenderResultsSections(IReadOnlyList<QueryResultSet> resultSets)
    {
        ResultsSectionsPanel.Children.Clear();

        for (var i = 0; i < resultSets.Count; i++)
        {
            var resultSet = resultSets[i];
            var expander = new Expander
            {
                Margin = new Thickness(0, 0, 0, 8),
                IsExpanded = i == 0,
                Header = resultSet.DataTable is null
                    ? $"{resultSet.Title} ({resultSet.AffectedRows} affected rows)"
                    : $"{resultSet.Title} ({resultSet.DataTable.Rows.Count} rows, {resultSet.DataTable.Columns.Count} columns)"
            };

            if (resultSet.DataTable is null)
            {
                expander.Content = new TextBlock
                {
                    Margin = new Thickness(8),
                    Text = $"No tabular rows. Affected rows: {resultSet.AffectedRows}."
                };
            }
            else
            {
                var dataGrid = CreateResultDataGrid(resultSet.DataTable);
                expander.Content = dataGrid;
            }

            ResultsSectionsPanel.Children.Add(expander);
        }
    }

    private DataGrid CreateResultDataGrid(DataTable table)
    {
        var dataGrid = new DataGrid
        {
            Margin = new Thickness(0, 8, 0, 0),
            MaxHeight = 400,
            IsReadOnly = true,
            AutoGenerateColumns = true,
            HeadersVisibility = DataGridHeadersVisibility.All,
            CanUserAddRows = false,
            CanUserDeleteRows = false,
            ItemsSource = table.DefaultView,
            Tag = table,
            FlowDirection = FlowDirection.LeftToRight,
            VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
            HorizontalScrollBarVisibility = ScrollBarVisibility.Auto
        };

        dataGrid.Loaded += ResultsDataGrid_Loaded;
        dataGrid.AutoGeneratingColumn += ResultsDataGrid_AutoGeneratingColumn;
        dataGrid.PreviewMouseRightButtonDown += ResultsDataGrid_PreviewMouseRightButtonDown;

        var contextMenu = new ContextMenu();
        var copyMenuItem = new MenuItem { Header = "Copy Value" };
        copyMenuItem.Click += CopyDataGridCellValueMenuItem_Click;
        contextMenu.Items.Add(copyMenuItem);

        var markCellMenuItem = new MenuItem { Header = "Mark This Cell" };
        markCellMenuItem.Click += MarkResultCellMenuItem_Click;
        contextMenu.Items.Add(markCellMenuItem);

        var markRowMenuItem = new MenuItem { Header = "Mark This Row" };
        markRowMenuItem.Click += MarkResultRowMenuItem_Click;
        contextMenu.Items.Add(markRowMenuItem);

        contextMenu.Items.Add(new Separator());
        var exportCsvMenuItem = new MenuItem { Header = "Export This Section as CSV..." };
        exportCsvMenuItem.Click += ExportResultSectionCsvMenuItem_Click;
        contextMenu.Items.Add(exportCsvMenuItem);

        var exportExcelMenuItem = new MenuItem { Header = "Export This Section as Excel..." };
        exportExcelMenuItem.Click += ExportResultSectionExcelMenuItem_Click;
        contextMenu.Items.Add(exportExcelMenuItem);

        dataGrid.ContextMenu = contextMenu;

        return dataGrid;
    }

    private void ResultsDataGrid_Loaded(object sender, RoutedEventArgs e)
    {
        if (sender is not DataGrid dataGrid || dataGrid.Items.Count == 0 || dataGrid.Columns.Count == 0)
        {
            return;
        }

        Dispatcher.BeginInvoke(DispatcherPriority.Background, new Action(() =>
        {
            if (dataGrid.Items.Count == 0 || dataGrid.Columns.Count == 0)
            {
                return;
            }

            var firstItem = dataGrid.Items[0];
            dataGrid.SelectedItem = firstItem;
            dataGrid.CurrentCell = new DataGridCellInfo(firstItem, dataGrid.Columns[0]);

            // Reset scroll position to show leftmost columns without using ScrollIntoView
            if (FindDescendant<ScrollViewer>(dataGrid) is ScrollViewer scrollViewer)
            {
                scrollViewer.ScrollToHorizontalOffset(0);
                scrollViewer.ScrollToVerticalOffset(0);
            }
        }));
    }

    private static string BuildResultsSummaryText(IReadOnlyList<QueryResultSet> resultSets, int fallbackAffectedRows)
    {
        var tableCount = resultSets.Count(x => x.DataTable is not null);
        if (tableCount == 0)
        {
            var affectedRows = resultSets.Sum(x => x.AffectedRows);
            var displayedAffectedRows = affectedRows == 0 ? fallbackAffectedRows : affectedRows;
            return $"Results ({displayedAffectedRows} affected rows)";
        }

        var totalRows = resultSets
            .Where(x => x.DataTable is not null)
            .Sum(x => x.DataTable!.Rows.Count);

        return $"Results ({tableCount} sections, {totalRows} total rows)";
    }

    private async void ExportResultSectionCsvMenuItem_Click(object sender, RoutedEventArgs e)
    {
        if (!TryGetResultSectionDataTableFromMenu(sender, out var table))
        {
            SetStatus("No result section selected for export.");
            return;
        }

        var dialog = new SaveFileDialog
        {
            Filter = "CSV Files (*.csv)|*.csv",
            FileName = "query-section.csv"
        };

        if (dialog.ShowDialog(this) != true)
        {
            return;
        }

        await _exportService.ExportToCsvAsync(table, dialog.FileName, _currentFullOutputMode, CancellationToken.None);
        SetStatus($"Section CSV export complete: {dialog.FileName}");
    }

    private async void ExportResultSectionExcelMenuItem_Click(object sender, RoutedEventArgs e)
    {
        if (!TryGetResultSectionDataTableFromMenu(sender, out var table))
        {
            SetStatus("No result section selected for export.");
            return;
        }

        var dialog = new SaveFileDialog
        {
            Filter = "Excel Files (*.xlsx)|*.xlsx",
            FileName = "query-section.xlsx"
        };

        if (dialog.ShowDialog(this) != true)
        {
            return;
        }

        await _exportService.ExportToExcelAsync(table, dialog.FileName, _currentFullOutputMode, CancellationToken.None);
        SetStatus($"Section Excel export complete: {dialog.FileName}");
    }

    private void MarkResultCellMenuItem_Click(object sender, RoutedEventArgs e)
    {
        if (sender is not MenuItem || e.OriginalSource is not MenuItem menuItem)
        {
            return;
        }

        var contextMenu = menuItem.Parent as ContextMenu;
        if (contextMenu?.PlacementTarget is not DataGrid dataGrid)
        {
            return;
        }

        var currentCell = dataGrid.CurrentCell;
        if (currentCell.Item is null || currentCell.Column is null)
        {
            SetStatus("Select a cell to mark.");
            return;
        }

        SetStatus($"Cell marked: [{currentCell.Column.Header}]");
    }

    private void MarkResultRowMenuItem_Click(object sender, RoutedEventArgs e)
    {
        if (sender is not MenuItem || e.OriginalSource is not MenuItem menuItem)
        {
            return;
        }

        var contextMenu = menuItem.Parent as ContextMenu;
        if (contextMenu?.PlacementTarget is not DataGrid dataGrid)
        {
            return;
        }

        if (dataGrid.SelectedItem is null)
        {
            SetStatus("Select a row to mark.");
            return;
        }

        var rowIndex = dataGrid.Items.IndexOf(dataGrid.SelectedItem) + 1;
        SetStatus($"Row marked: #{rowIndex}");
    }

    private static bool TryGetResultSectionDataTableFromMenu(object sender, out DataTable table)
    {
        table = null!;

        if (sender is not MenuItem menuItem
            || menuItem.Parent is not ContextMenu contextMenu
            || contextMenu.PlacementTarget is not DataGrid dataGrid
            || dataGrid.Tag is not DataTable dataTable)
        {
            return false;
        }

        table = dataTable;
        return true;
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

    private async void SchemaDataGrid_PreviewMouseLeftButtonUp(object sender, MouseButtonEventArgs e)
    {
        // Double-click copy behavior is now completely disabled in all tabs
        e.Handled = false;
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

    private async void SchemaListBox_PreviewMouseLeftButtonUp(object sender, MouseButtonEventArgs e)
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

        if (await TrySetClipboardTextAsync(copiedText))
        {
            SetStatus($"Copied value to clipboard: {TruncateForStatus(copiedText)}");
        }
        else
        {
            SetStatus("Could not access the clipboard. Please try again.");
        }
    }

    private async void CopySelectedObjectNameMenuItem_Click(object sender, RoutedEventArgs e)
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

        if (await TrySetClipboardTextAsync(selectedName))
        {
            SetStatus($"Copied value to clipboard: {selectedName}");
        }
        else
        {
            SetStatus("Could not access the clipboard. Please try again.");
        }
    }

    private static async Task<bool> TrySetClipboardTextAsync(string text)
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
                await Task.Delay(retryDelayMs);
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

        return TryGetDataGridCellTextForCopy(item, column)
            ?? TryGetCellContentTextForCopy(cell.Content);
    }

    private bool TryGetCurrentDataGridCellTextForCopy(DataGrid dataGrid, out string copiedText)
    {
        copiedText = string.Empty;

        var currentCell = dataGrid.CurrentCell;
        var item = currentCell.Item;
        var column = currentCell.Column;

        var text = TryGetDataGridCellTextForCopy(item, column);
        if (string.IsNullOrWhiteSpace(text) && dataGrid.SelectedItem is not null && column is not null)
        {
            text = TryGetDataGridCellTextForCopy(dataGrid.SelectedItem, column);
        }

        if (string.IsNullOrWhiteSpace(text))
        {
            return false;
        }

        copiedText = text;
        return true;
    }

    private string? TryGetDataGridCellTextForCopy(object? item, DataGridColumn? column)
    {
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

        return null;
    }

    private static string? TryGetCellContentTextForCopy(object? content)
    {
        if (content is TextBlock textBlock)
        {
            return textBlock.Text;
        }

        if (content is CheckBox checkBox)
        {
            return checkBox.IsChecked?.ToString();
        }

        if (content is TextBox textBox)
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
            if (sender == EditRowsDataGrid && e.Column is DataGridCheckBoxColumn checkBoxColumn)
            {
                if (checkBoxColumn.Binding is Binding checkBoxBinding)
                {
                    checkBoxBinding.Mode = BindingMode.TwoWay;
                    checkBoxBinding.UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged;
                }
            }

            return;
        }

        if (textColumn.Binding is not Binding binding)
        {
            return;
        }

        binding.Converter = ResultsValueConverter;

        if (sender == EditRowsDataGrid)
        {
            var editingElementStyle = new Style(typeof(TextBox));
            editingElementStyle.Setters.Add(new Setter(TextBox.BackgroundProperty, FindResource("InputBackgroundBrush")));
            editingElementStyle.Setters.Add(new Setter(TextBox.ForegroundProperty, FindResource("InputForegroundBrush")));
            editingElementStyle.Setters.Add(new Setter(TextBox.BorderBrushProperty, FindResource("AccentBrush")));
            editingElementStyle.Setters.Add(new Setter(TextBox.BorderThicknessProperty, new Thickness(1)));
            textColumn.EditingElementStyle = editingElementStyle;
        }
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

    private static T? FindDescendant<T>(DependencyObject current) where T : DependencyObject
    {
        var childrenCount = VisualTreeHelper.GetChildrenCount(current);
        for (var i = 0; i < childrenCount; i++)
        {
            var child = VisualTreeHelper.GetChild(current, i);
            if (child is T match)
            {
                return match;
            }

            var foundInChild = FindDescendant<T>(child);
            if (foundInChild is not null)
            {
                return foundInChild;
            }
        }

        return null;
    }

    private static T? FindDescendantByName<T>(DependencyObject current, string targetName) where T : FrameworkElement
    {
        var childrenCount = VisualTreeHelper.GetChildrenCount(current);
        for (var i = 0; i < childrenCount; i++)
        {
            var child = VisualTreeHelper.GetChild(current, i);
            if (child is T typed && string.Equals(typed.Name, targetName, StringComparison.Ordinal))
            {
                return typed;
            }

            var foundInChild = FindDescendantByName<T>(child, targetName);
            if (foundInChild is not null)
            {
                return foundInChild;
            }
        }

        return null;
    }

    private void ApplyEditRowsCornerButtonStyle()
    {
        Dispatcher.BeginInvoke(DispatcherPriority.Loaded, new Action(() =>
        {
            EditRowsDataGrid.ApplyTemplate();
            var selectAllButton = FindDescendantByName<Button>(EditRowsDataGrid, "PART_SelectAllButton");
            if (selectAllButton is null)
            {
                return;
            }

            selectAllButton.SetResourceReference(Control.BackgroundProperty, "PanelBrush");
            selectAllButton.SetResourceReference(Control.ForegroundProperty, "TextSecondaryBrush");
            selectAllButton.SetResourceReference(Control.BorderBrushProperty, "BorderBrush");
            selectAllButton.BorderThickness = new Thickness(0, 0, 1, 1);
        }));
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

    private void ExitEditModeAndClearEditableRows()
    {
        _isEditMode = false;
        _editableResultsTable = null;
        ApplyEditModeState();
        RefreshEditRowsVisualStates();
    }

    private void ShowSelectedTableColumnsInEditRowsGrid()
    {
        ExitEditModeAndClearEditableRows();

        if (_selectedTable is null || _selectedColumns.Count == 0)
        {
            EditRowsDataGrid.ItemsSource = null;
            return;
        }

        var schemaPreviewTable = new DataTable();
        foreach (var column in _selectedColumns.OrderBy(x => x.OrdinalPosition))
        {
            schemaPreviewTable.Columns.Add(column.ColumnName, typeof(object));
        }

        EditRowsDataGrid.ItemsSource = schemaPreviewTable.DefaultView;
        EditRowsSummaryTextBlock.Text = $"Selected table: {_selectedTable.FullName} ({_selectedColumns.Count} columns). Click Load to fetch rows.";
        RefreshEditRowsVisualStates();
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

    private void RefreshEditRowsVisualStates()
    {
        if (EditRowsDataGrid.Items.Count == 0)
        {
            return;
        }

        for (var i = 0; i < EditRowsDataGrid.Items.Count; i++)
        {
            var row = EditRowsDataGrid.ItemContainerGenerator.ContainerFromIndex(i) as DataGridRow;
            if (row is null)
            {
                continue;
            }

            ApplyEditRowsRowVisualState(row);
        }
    }

    private void ApplyEditRowsVisualStateAtIndex(int index)
    {
        if (index < 0 || index >= EditRowsDataGrid.Items.Count)
        {
            return;
        }

        var row = EditRowsDataGrid.ItemContainerGenerator.ContainerFromIndex(index) as DataGridRow;
        if (row is null)
        {
            return;
        }

        ApplyEditRowsRowVisualState(row);
    }

    private void ApplyEditRowsRowVisualState(DataGridRow row)
    {
        row.Tag = null;

        if (row.Item is not DataRowView rowView)
        {
            row.Header = row.GetIndex() + 1;
            return;
        }

        switch (rowView.Row.RowState)
        {
            case DataRowState.Added:
                row.Tag = "Added";
                row.Header = "+";
                break;
            case DataRowState.Modified:
                row.Tag = "Modified";
                row.Header = "*";
                break;
            case DataRowState.Deleted:
                row.Tag = "Deleted";
                row.Header = "-";
                break;
            default:
                row.Header = row.GetIndex() + 1;
                break;
        }
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

        SetExecutionState(true);
        SetStatus(_isEditRowsCustomQueryMode
            ? (_currentFullOutputMode
                ? $"Loading editable rows from custom SQL (full output mode)..."
                : "Loading editable rows from custom SQL...")
            : (_currentFullOutputMode
                ? $"Loading editable rows from {_selectedTable.FullName} (full output mode)..."
                : $"Loading editable rows from {_selectedTable.FullName}..."));

        try
        {
            if (_isEditRowsCustomQueryMode)
            {
                var sqlToExecute = editQueryMode.Sql;
                if (string.IsNullOrWhiteSpace(sqlToExecute))
                {
                    SetStatus("Edit query is required.");
                    return;
                }

                TrackRecentSqlFragments(sqlToExecute);

                var result = await _databaseQueryService.ExecuteAsync(
                    connectionString,
                    sqlToExecute,
                    ParseTimeoutSeconds(),
                    CancellationToken.None);

                if (!result.IsSuccess)
                {
                    SetStatus($"Failed to load editable rows: {result.ErrorMessage}");
                    return;
                }

                if (result.DataTable is null)
                {
                    SetStatus("Custom SQL did not return tabular rows.");
                    return;
                }

                _editableResultsTable = result.DataTable;
                if (_editableResultsTable is not null)
                {
                    _editableResultsTable.AcceptChanges();
                }
            }
            else
            {
                var topRows = ParseEditTopRowsInput();
                var filter = string.IsNullOrWhiteSpace(EditFilterTextBox.Text) ? null : EditFilterTextBox.Text.Trim();
                var orderBy = string.IsNullOrWhiteSpace(EditOrderByTextBox.Text) ? null : EditOrderByTextBox.Text.Trim();

                if (topRows > 5000)
                {
                    topRows = 5000;
                    _isSyncingEditQuery = true;
                    EditTopRowsTextBox.Text = "5000";
                    _isSyncingEditQuery = false;
                    UpdateEditQueryTextFromInputs();
                }

                _editableResultsTable = await _rowEditService.LoadTopRowsAsync(
                    connectionString,
                    _selectedTable.SchemaName,
                    _selectedTable.TableName,
                    topRows,
                    filter,
                    orderBy,
                    ParseTimeoutSeconds(),
                    CancellationToken.None);
            }

            if (_editableResultsTable is null)
            {
                SetStatus("No editable rows were returned.");
                return;
            }

            _isEditMode = true;
            _currentDataTable = _editableResultsTable;
            EditRowsDataGrid.ItemsSource = _editableResultsTable.DefaultView;
            EditRowsSummaryTextBlock.Text = _isEditRowsCustomQueryMode
                ? $"Edit Mode (Custom SQL, {_editableResultsTable.Rows.Count} rows loaded)"
                : $"Edit Mode ({_editableResultsTable.Rows.Count} rows loaded)";
            OutputTabControl.SelectedIndex = OutputEditRowsTabIndex;
            ApplyEditModeState();
            RefreshEditRowsVisualStates();

            // SetStatus(_isEditRowsCustomQueryMode
            // ? "Edit mode ready from custom SQL."
            // : $"Edit mode ready for {_selectedTable.FullName}.");
            
            SetStatus($"({_editableResultsTable.Rows.Count}) rows loaded.");
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
        if (_isSyncingEditQuery || _isEditRowsCustomQueryMode)
        {
            return;
        }

        UpdateEditQueryTextFromInputs();
    }

    private void EditQueryTextBox_TextChanged(object sender, EventArgs e)
    {
        if (_isSyncingEditQuery)
        {
            return;
        }

        if (!_isEditRowsCustomQueryMode)
        {
            SetEditRowsQueryMode(true);
            _ = UpdateSqlSuggestionsForAsync(_editRowsSqlEditor!);
            return;
        }

        SyncInputsFromEditQueryText();
        _ = UpdateSqlSuggestionsForAsync(_editRowsSqlEditor!);
    }

    private void SqlEditorTextBox_TextChanged(object sender, EventArgs e)
    {
        if (ReferenceEquals(sender, QueryTextBox))
        {
            _ = UpdateSqlSuggestionsForAsync(_sqlEditor!);
            return;
        }

        if (ReferenceEquals(sender, EditQueryTextBox))
        {
            _ = UpdateSqlSuggestionsForAsync(_editRowsSqlEditor!);
        }
    }

    private void SqlEditorTextBox_PreviewKeyDown(object sender, KeyEventArgs e)
    {
        var editor = ReferenceEquals(sender, QueryTextBox)
            ? _sqlEditor
            : ReferenceEquals(sender, EditQueryTextBox)
                ? _editRowsSqlEditor
                : null;

        if (editor is null)
        {
            return;
        }

        if (Keyboard.Modifiers == ModifierKeys.Control && e.Key == Key.Space)
        {
            _ = UpdateSqlSuggestionsForAsync(editor, allowEmptyPrefix: true, debounceMs: 0);
            e.Handled = true;
            return;
        }

        if (!SqlSuggestionPopup.IsOpen || !ReferenceEquals(_activeSqlSuggestionTextEditor, editor))
        {
            return;
        }

        if (e.Key == Key.Down)
        {
            if (SqlSuggestionListBox.Items.Count == 0)
            {
                return;
            }

            var nextIndex = Math.Min(SqlSuggestionListBox.SelectedIndex + 1, SqlSuggestionListBox.Items.Count - 1);
            SqlSuggestionListBox.SelectedIndex = Math.Max(0, nextIndex);
            SqlSuggestionListBox.ScrollIntoView(SqlSuggestionListBox.SelectedItem);
            e.Handled = true;
            return;
        }

        if (e.Key == Key.Up)
        {
            if (SqlSuggestionListBox.Items.Count == 0)
            {
                return;
            }

            var previousIndex = Math.Max(SqlSuggestionListBox.SelectedIndex - 1, 0);
            SqlSuggestionListBox.SelectedIndex = previousIndex;
            SqlSuggestionListBox.ScrollIntoView(SqlSuggestionListBox.SelectedItem);
            e.Handled = true;
            return;
        }

        if (e.Key is Key.Enter or Key.Tab)
        {
            if (TryApplySelectedSqlSuggestion())
            {
                e.Handled = true;
            }

            return;
        }

        if (e.Key == Key.Escape)
        {
            HideSqlSuggestions();
            e.Handled = true;
        }
    }

    private void SqlSuggestionListBox_MouseDoubleClick(object sender, MouseButtonEventArgs e)
    {
        TryApplySelectedSqlSuggestion();
    }

    private async Task UpdateSqlSuggestionsForAsync(ISqlTextEditor editor, bool allowEmptyPrefix = false, int debounceMs = 120)
    {
        if (editor is null)
        {
            HideSqlSuggestions();
            return;
        }

        _sqlSuggestionDebounceCts?.Cancel();
        _sqlSuggestionDebounceCts?.Dispose();
        _sqlSuggestionDebounceCts = new CancellationTokenSource();
        var cancellationToken = _sqlSuggestionDebounceCts.Token;

        try
        {
            if (debounceMs > 0)
            {
                await Task.Delay(debounceMs, cancellationToken);
            }
        }
        catch (OperationCanceledException)
        {
            return;
        }

        if (cancellationToken.IsCancellationRequested)
        {
            return;
        }

        if (!TryGetCurrentSqlToken(editor, out var token, out var tokenStart, out var tokenLength))
        {
            HideSqlSuggestions();
            return;
        }

        if (!allowEmptyPrefix && string.IsNullOrWhiteSpace(token))
        {
            HideSqlSuggestions();
            return;
        }

        var catalogSnapshot = _sqlCompletionCatalogService.GetSnapshot();

        var tableCandidates = catalogSnapshot.Tables.Count > 0
            ? catalogSnapshot.Tables
            : _tables;

        var storedProcedureCandidates = catalogSnapshot.StoredProcedures.Count > 0
            ? catalogSnapshot.StoredProcedures
            : _storedProcedures;

        var columnCandidates = catalogSnapshot.GlobalColumns
            .Concat(_selectedColumns)
            .GroupBy(x => $"{x.SchemaName}.{x.TableName}.{x.ColumnName}", StringComparer.OrdinalIgnoreCase)
            .Select(x => x.First())
            .ToList();

        var procedureParameterCandidates = catalogSnapshot.ProcedureParameters
            .Concat(_selectedProcedureParameters)
            .GroupBy(x => x.ParameterName, StringComparer.OrdinalIgnoreCase)
            .Select(x => x.First())
            .ToList();

        var request = new SqlSuggestionRequest
        {
            SqlText = editor.Text ?? string.Empty,
            Token = token,
            TokenStart = tokenStart,
            Keywords = catalogSnapshot.Keywords,
            Functions = catalogSnapshot.Functions,
            Snippets = catalogSnapshot.Snippets,
            Tables = tableCandidates,
            SelectedColumns = columnCandidates,
            StoredProcedures = storedProcedureCandidates,
            ProcedureParameters = procedureParameterCandidates,
            ForeignKeys = catalogSnapshot.ForeignKeys.Count > 0 ? catalogSnapshot.ForeignKeys : _foreignKeys,
            RecentFragments = GetRecentSqlFragments(),
            MaxResults = 30
        };

        var suggestions = _sqlSuggestionEngine.GetSuggestions(request);
        if (suggestions.Count == 0)
        {
            HideSqlSuggestions();
            return;
        }

        _activeSqlSuggestionTextEditor = editor;
        _activeSqlSuggestionTokenStart = tokenStart;
        _activeSqlSuggestionTokenLength = tokenLength;

        SqlSuggestionListBox.ItemsSource = suggestions;
        SqlSuggestionListBox.SelectedIndex = 0;

        SqlSuggestionPopup.PlacementTarget = editor.Element;
        SqlSuggestionPopup.HorizontalOffset = 16;
        SqlSuggestionPopup.VerticalOffset = 24;
        SqlSuggestionPopup.IsOpen = true;
    }

    private static bool TryGetCurrentSqlToken(ISqlTextEditor editor, out string token, out int tokenStart, out int tokenLength)
    {
        token = string.Empty;
        tokenStart = 0;
        tokenLength = 0;

        if (editor is null)
        {
            return false;
        }

        var text = editor.Text ?? string.Empty;
        var caret = Math.Clamp(editor.CaretIndex, 0, text.Length);

        var start = caret;
        while (start > 0 && IsSqlTokenCharacter(text[start - 1]))
        {
            start--;
        }

        var end = caret;
        while (end < text.Length && IsSqlTokenCharacter(text[end]))
        {
            end++;
        }

        tokenStart = start;
        tokenLength = end - start;
        if (tokenLength < 0)
        {
            return false;
        }

        token = tokenLength == 0 ? string.Empty : text.Substring(tokenStart, tokenLength);
        return true;
    }

    private static bool IsSqlTokenCharacter(char value)
    {
        return char.IsLetterOrDigit(value)
            || value is '_' or '@' or '[' or ']' or '.';
    }

    private bool TryApplySelectedSqlSuggestion()
    {
        if (!SqlSuggestionPopup.IsOpen
            || _activeSqlSuggestionTextEditor is null
            || SqlSuggestionListBox.SelectedItem is not string selectedSuggestion)
        {
            return false;
        }

        var editor = _activeSqlSuggestionTextEditor;
        var safeStart = Math.Clamp(_activeSqlSuggestionTokenStart, 0, editor.Text.Length);
        var safeLength = Math.Clamp(_activeSqlSuggestionTokenLength, 0, editor.Text.Length - safeStart);

        editor.Select(safeStart, safeLength);
        editor.ReplaceSelection(selectedSuggestion + " ");
        editor.CaretIndex = safeStart + selectedSuggestion.Length + 1;
        editor.Select(editor.CaretIndex, 0);
        editor.Focus();

        HideSqlSuggestions();
        return true;
    }

    private void HideSqlSuggestions()
    {
        SqlSuggestionPopup.IsOpen = false;
        SqlSuggestionListBox.ItemsSource = null;
        _activeSqlSuggestionTextEditor = null;
        _activeSqlSuggestionTokenStart = 0;
        _activeSqlSuggestionTokenLength = 0;
    }

    private void TrackRecentSqlFragments(string sqlText)
    {
        if (string.IsNullOrWhiteSpace(sqlText))
        {
            return;
        }

        foreach (var fragment in ExtractSqlFragments(sqlText))
        {
            if (_recentSqlFragmentLookup.Contains(fragment))
            {
                continue;
            }

            _recentSqlFragments.AddFirst(fragment);
            _recentSqlFragmentLookup.Add(fragment);

            while (_recentSqlFragments.Count > MaxRecentSqlFragments)
            {
                var last = _recentSqlFragments.Last;
                if (last is null)
                {
                    break;
                }

                _recentSqlFragmentLookup.Remove(last.Value);
                _recentSqlFragments.RemoveLast();
            }
        }
    }

    private IReadOnlyList<string> GetRecentSqlFragments()
    {
        return _recentSqlFragments.ToList();
    }

    private static IEnumerable<string> ExtractSqlFragments(string sqlText)
    {
        return sqlText
            .Split(';', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
            .Select(static part => Regex.Replace(part, @"\s+", " ").Trim())
            .Where(static part => part.Length >= 12)
            .Select(static part => part.Length > 120 ? part[..120] + "..." : part)
            .Take(6);
    }

    private void UpdateEditQueryTextFromInputs()
    {
        if (_isSyncingEditQuery || _isEditRowsCustomQueryMode)
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
        EditQueryTextBox.CaretOffset = EditQueryTextBox.Text.Length;
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

        if (match.Groups["top"].Success
            && (!int.TryParse(match.Groups["top"].Value, out topRows) || topRows <= 0))
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

    private void EditRowsCustomQueryModeCheckBox_Checked(object sender, RoutedEventArgs e)
    {
        SetEditRowsQueryMode(true, updateToggle: false);
    }

    private void EditRowsCustomQueryModeCheckBox_Unchecked(object sender, RoutedEventArgs e)
    {
        SetEditRowsQueryMode(false, updateToggle: false);
    }

    private void SetEditRowsQueryMode(bool customMode, bool updateToggle = true)
    {
        if (_isEditRowsCustomQueryMode == customMode)
        {
            if (updateToggle && EditRowsCustomQueryModeCheckBox is not null)
            {
                _isSyncingEditQuery = true;
                EditRowsCustomQueryModeCheckBox.IsChecked = customMode;
                _isSyncingEditQuery = false;
            }

            return;
        }

        if (!customMode)
        {
            var generatedSql = BuildEditRowsQuery(
                ParseEditTopRowsInput(),
                string.IsNullOrWhiteSpace(EditFilterTextBox.Text) ? null : EditFilterTextBox.Text.Trim(),
                string.IsNullOrWhiteSpace(EditOrderByTextBox.Text) ? null : EditOrderByTextBox.Text.Trim());

            if (!string.Equals(EditQueryTextBox.Text, generatedSql, StringComparison.Ordinal))
            {
                var decision = MessageBox.Show(
                    "Switching to structured mode will replace the current custom query text with a generated query based on Top/Filter/Order. Continue?",
                    "Switch Query Mode",
                    MessageBoxButton.YesNo,
                    MessageBoxImage.Question);

                if (decision != MessageBoxResult.Yes)
                {
                    if (EditRowsCustomQueryModeCheckBox is not null)
                    {
                        _isSyncingEditQuery = true;
                        EditRowsCustomQueryModeCheckBox.IsChecked = true;
                        _isSyncingEditQuery = false;
                    }

                    return;
                }
            }
        }

        _isEditRowsCustomQueryMode = customMode;

        if (updateToggle && EditRowsCustomQueryModeCheckBox is not null)
        {
            _isSyncingEditQuery = true;
            EditRowsCustomQueryModeCheckBox.IsChecked = customMode;
            _isSyncingEditQuery = false;
        }

        if (EditTopRowsTextBox is not null)
        {
            EditTopRowsTextBox.IsEnabled = !customMode;
        }

        if (EditFilterTextBox is not null)
        {
            EditFilterTextBox.IsEnabled = !customMode;
        }

        if (EditOrderByTextBox is not null)
        {
            EditOrderByTextBox.IsEnabled = !customMode;
        }

        if (customMode)
        {
            SetStatus("Custom SQL mode enabled for Edit Rows.");
            SyncInputsFromEditQueryText();
            return;
        }

        SetStatus("Structured mode enabled for Edit Rows.");
        UpdateEditQueryTextFromInputs();
    }

    private static IReadOnlyList<RowUpdateRequest> BuildRowUpdates(DataTable table, IReadOnlyList<ColumnSchemaInfo> columns)
    {
        var columnNames = columns.Select(c => c.ColumnName).ToHashSet(StringComparer.OrdinalIgnoreCase);

        var updates = new List<RowUpdateRequest>();
        foreach (DataRow row in table.Rows)
        {
            if (row.RowState != DataRowState.Modified)
            {
                continue;
            }

            var keyValues = new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase);
            foreach (DataColumn column in table.Columns)
            {
                if (!columnNames.Contains(column.ColumnName))
                {
                    continue;
                }

                var value = row[column, DataRowVersion.Original];
                keyValues[column.ColumnName] = value == DBNull.Value ? null : value;
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

    private sealed class DeleteColumnSelectionRow
    {
        public required string ColumnName { get; init; }

        public required string DataType { get; init; }

        public bool IsSelected { get; set; }
    }

    private sealed class QueryParameterEditorRow
    {
        public required string ParameterName { get; init; }

        public string? Value { get; set; }

        public bool SendAsNull { get; set; }
    }
}