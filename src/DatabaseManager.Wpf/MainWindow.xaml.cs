using System.Data;
using System.Collections.ObjectModel;
using System.Configuration;
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
using CommunityToolkit.Mvvm.Input;
using DatabaseManager.Wpf.Commands;
using DatabaseManager.Wpf.Editors;
using DatabaseManager.Wpf.SqlSuggestions;
using DatabaseManager.Wpf.ViewModels;
using DatabaseManager.Wpf.Windows;
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
    private readonly IConnectionProfileStoreService _connectionProfileStoreService = new ConnectionProfileStoreService();
    private static readonly IValueConverter ResultsValueConverter = new ResultValueConverter();

    private readonly ICommandRegistry _commandRegistry = new AppCommandRegistry();
    private readonly IToastService _toastService;
    public MainWindowViewModel ViewModel { get; }

    private Guid? _selectedConnectionProfileId;
    private string? _selectedConnectionProfileName;

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
    private bool _isSettingQueryTextProgrammatically;
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
        var defaultConnectionString = ConfigurationManager.AppSettings["DefaultConnectionString"];
        if (!string.IsNullOrWhiteSpace(defaultConnectionString))
        {
            ConnectionStringTextBox.Text = defaultConnectionString;
        }
        UpdateConnectionSummaryDisplay();
        _sqlEditor = new AvalonEditSqlTextEditorAdapter(QueryTextBox);
        _editRowsSqlEditor = new AvalonEditSqlTextEditorAdapter(EditQueryTextBox);
        SqlEditorSupport.Configure(QueryTextBox);
        SqlEditorSupport.Configure(EditQueryTextBox);
        _toastService = new ToastService(Dispatcher);
        ToastHostControl.Toasts = _toastService.Toasts;
        var templatesPanel = new TemplatesPanelViewModel(
            _templateStoreService,
            getCurrentSqlText: () => QueryTextBox.Text,
            onTemplateActivated: OnTemplateActivated,
            confirmDelete: ConfirmDeleteTemplate,
            setStatus: SetStatus);
        var schemaAssistant = new SchemaAssistantViewModel(
            _databaseSchemaService,
            _queryAssistantService,
            _sqlCompletionCatalogService,
            getConnectionString: () => ConnectionStringTextBox.Text.Trim(),
            onTableSelected: OnSchemaTableSelected,
            onTableCleared: OnSchemaTableCleared,
            onProcedureSelected: OnSchemaProcedureSelected,
            onProcedureCleared: OnSchemaProcedureCleared,
            onSchemaMetadataLoaded: OnSchemaMetadataLoaded,
            onScriptGenerated: OnSchemaScriptGenerated,
            onOpenInRunnerRequested: OnOpenInRunnerRequested,
            onCopyRequested: CopySchemaObjectNameToClipboardAsync,
            setStatus: SetStatus);
        ViewModel = new MainWindowViewModel(_commandRegistry, OnDarkModeChanged, OnSchemaAssistantVisibleChanged, templatesPanel, schemaAssistant);
        DataContext = ViewModel;
        RegisterCommands();
        BuildInputBindings();
        SourceInitialized += MainWindow_SourceInitialized;
        Loaded += MainWindow_Loaded;
    }

    [DllImport("dwmapi.dll")]
    private static extern int DwmSetWindowAttribute(IntPtr hwnd, int dwAttribute, ref int pvAttribute, int cbAttribute);

    private void MainWindow_SourceInitialized(object? sender, EventArgs e)
    {
        ApplyTitleBarTheme(ViewModel.IsDarkMode);
    }

    private async void MainWindow_Loaded(object sender, RoutedEventArgs e)
    {
        App.ApplyTheme(ViewModel.IsDarkMode);
        ApplyTitleBarTheme(ViewModel.IsDarkMode);
        await ViewModel.TemplatesPanel.RefreshAsync();
        RunnerParametersDataGrid.ItemsSource = _runnerParameterRows;
        OutputTabControl.SelectedIndex = OutputEditRowsTabIndex;
        ApplyEditModeState();
        ApplyEditRowsCornerButtonStyle();
        UpdateEditQueryTextFromInputs();
        await ApplyDefaultConnectionProfileAsync();
        await ConnectToDatabaseAsync(triggeredOnStartup: true);
        if (_tables.Count == 0 && _storedProcedures.Count == 0)
        {
            SetStatus("Ready. Shortcuts: Ctrl+E to run query, Ctrl+Q to cancel.");
        }
    }

    private void OnSchemaAssistantVisibleChanged(bool visible)
    {
        if (visible)
        {
            SchemaAssistantColumn.Width = _schemaPaneExpandedWidth;
            SchemaAssistantSplitterColumn.Width = new GridLength(4);
            SchemaAssistantPanel.Visibility = Visibility.Visible;
            SchemaPanelSplitter.Visibility = Visibility.Visible;
        }
        else
        {
            SchemaAssistantColumn.Width = new GridLength(0);
            SchemaAssistantSplitterColumn.Width = new GridLength(0);
            SchemaAssistantPanel.Visibility = Visibility.Collapsed;
            SchemaPanelSplitter.Visibility = Visibility.Collapsed;
        }
    }

    private void OnDarkModeChanged(bool isDarkMode)
    {
        App.ApplyTheme(isDarkMode);
        ApplyTitleBarTheme(isDarkMode);
        ApplyEditRowsCornerButtonStyle();
    }

    /// <summary>
    /// Registers every app-level command once, shared by the top menu, the command palette,
    /// the keyboard shortcuts panel, and (via <see cref="BuildInputBindings"/>) the window's
    /// InputBindings - see CLAUDE.md / the command-registry design note. Contextual behavior
    /// that used to live in the old MainWindow_PreviewKeyDown switch (e.g. Ctrl+R/S only
    /// applying while the Edit Rows tab is active) is preserved by checking
    /// OutputTabControl.SelectedIndex inside the command body itself, exactly like the
    /// original imperative code did.
    /// </summary>
    private void RegisterCommands()
    {
        void Reg(string id, string displayName, string category, KeyGesture? gesture, ICommand command, string? iconKey = null, params string[] keywords)
        {
            _commandRegistry.Register(new AppCommandDescriptor
            {
                Id = id,
                DisplayName = displayName,
                Category = category,
                IconKey = iconKey,
                Gesture = gesture,
                Command = command,
                Keywords = keywords
            });
        }

        void RegisterTabSwitch(string id, string displayName, Key key, int tabIndex)
        {
            Reg(id, displayName, "View", new KeyGesture(key, ModifierKeys.Control),
                new RelayCommand(() => OutputTabControl.SelectedIndex = tabIndex), "Icon.Table");
        }

        Reg("connection.connect", "Connect", "Connection", null,
            new AsyncRelayCommand(() => ConnectToDatabaseAsync(triggeredOnStartup: false)), "Icon.Power");

        Reg("connection.manage", "Manage Connections...", "Connection", null,
            new RelayCommand(() => ConnectionSummaryButton_Click(ConnectionSummaryButton, new RoutedEventArgs())), "Icon.Power", "profile", "saved");

        Reg("query.run", "Run Query", "Query", new KeyGesture(Key.E, ModifierKeys.Control),
            new RelayCommand(() =>
            {
                if (OutputTabControl.SelectedIndex == OutputEditRowsTabIndex)
                {
                    _ = RefreshEditRowsAsync();
                }
                else
                {
                    RunQueryButton_Click(RunQueryButton, new RoutedEventArgs());
                }
            }), "Icon.Run");

        Reg("query.cancel", "Cancel Execution", "Query", new KeyGesture(Key.Q, ModifierKeys.Control),
            new RelayCommand(() => CancelButton_Click(CancelButton, new RoutedEventArgs())), "Icon.Stop");

        Reg("editRows.refresh", "Refresh Edit Rows", "Edit Rows", new KeyGesture(Key.R, ModifierKeys.Control),
            new AsyncRelayCommand(() => OutputTabControl.SelectedIndex == OutputEditRowsTabIndex
                ? RefreshEditRowsAsync()
                : Task.CompletedTask), "Icon.Refresh");

        Reg("editRows.save", "Save Row Changes", "Edit Rows", new KeyGesture(Key.S, ModifierKeys.Control),
            new AsyncRelayCommand(() => OutputTabControl.SelectedIndex == OutputEditRowsTabIndex
                ? SaveRowChangesAsync()
                : Task.CompletedTask), "Icon.Save");

        RegisterTabSwitch("view.switchEditRows", "Switch to Edit Rows", Key.D1, OutputEditRowsTabIndex);
        RegisterTabSwitch("view.switchSqlEditor", "Switch to SQL Editor", Key.D2, OutputSqlEditorTabIndex);
        RegisterTabSwitch("view.switchSchema", "Switch to Schema", Key.D3, OutputSchemaTabIndex);
        RegisterTabSwitch("view.switchResults", "Switch to Results", Key.D4, OutputResultsTabIndex);
        RegisterTabSwitch("view.switchProcedureRunner", "Switch to Procedure Runner", Key.D5, OutputProcedureRunnerTabIndex);

        Reg("view.toggleDarkMode", "Toggle Dark Mode", "View", null,
            new RelayCommand(() => ViewModel.IsDarkMode = !ViewModel.IsDarkMode), "Icon.Moon");

        Reg("view.toggleObjectExplorer", "Toggle Object Explorer", "View", null,
            new RelayCommand(() => ViewModel.IsSchemaAssistantVisible = !ViewModel.IsSchemaAssistantVisible), "Icon.Sidebar");

        Reg("view.openCommandPalette", "Command Palette", "View", new KeyGesture(Key.P, ModifierKeys.Control | ModifierKeys.Shift),
            new RelayCommand(OpenCommandPalette), "Icon.Palette");

        Reg("view.showShortcuts", "Keyboard Shortcuts", "View", new KeyGesture(Key.OemQuestion, ModifierKeys.Control),
            new RelayCommand(OpenShortcutsHelp), "Icon.Keyboard", "help");
    }

    private void BuildInputBindings()
    {
        foreach (var descriptor in _commandRegistry.Commands)
        {
            if (descriptor.Gesture is not null)
            {
                InputBindings.Add(new KeyBinding(descriptor.Command, descriptor.Gesture));
            }
        }
    }

    private void OpenCommandPalette()
    {
        var palette = new CommandPaletteWindow(_commandRegistry.Commands) { Owner = this };
        palette.ShowDialog();
    }

    private void OpenShortcutsHelp()
    {
        var help = new ShortcutsHelpWindow(_commandRegistry.Commands) { Owner = this };
        help.ShowDialog();
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

    private async void ConnectionSummaryButton_Click(object sender, RoutedEventArgs e)
    {
        var picker = new ConnectionPickerWindow(_connectionProfileStoreService, ConnectionStringTextBox.Text) { Owner = this };
        if (picker.ShowDialog() != true || picker.SelectedConnectionString is null)
        {
            return;
        }

        ConnectionStringTextBox.Text = picker.SelectedConnectionString;
        _selectedConnectionProfileId = picker.SelectedProfileId;
        _selectedConnectionProfileName = picker.SelectedProfileName;
        UpdateConnectionSummaryDisplay();

        // Picking a connection in the dialog should actually connect, not just fill in the
        // (now-hidden) textbox and leave the user to separately press the toolbar Connect
        // button - besides being the behavior a "Connect" button implies, TouchLastUsedAsync
        // only ever runs from inside ConnectToDatabaseAsync's success path, so without this
        // a profile's "last used" timestamp would never update from picking it alone.
        await ConnectToDatabaseAsync(triggeredOnStartup: false);
    }

    /// <summary>
    /// A saved connection profile marked "default" should be what the app auto-connects to
    /// on startup - that's the whole point of marking one default. Previously "default" only
    /// affected which profile the picker dialog pre-selected; startup only ever looked at
    /// App.config's DefaultConnectionString, with no awareness of profiles at all. Profiles
    /// now take priority; App.config's value remains the fallback when no profile is marked
    /// default (e.g. before any profiles have been created).
    /// </summary>
    private async Task ApplyDefaultConnectionProfileAsync()
    {
        var profiles = await _connectionProfileStoreService.GetAllAsync(CancellationToken.None);
        var defaultProfile = profiles.FirstOrDefault(p => p.IsDefault);
        if (defaultProfile is null)
        {
            return;
        }

        ConnectionStringTextBox.Text = _connectionProfileStoreService.Decrypt(defaultProfile);
        _selectedConnectionProfileId = defaultProfile.Id;
        _selectedConnectionProfileName = defaultProfile.Name;
        UpdateConnectionSummaryDisplay();
    }

    private void UpdateConnectionSummaryDisplay()
    {
        ConnectionSummaryTextBlock.Text = _selectedConnectionProfileName is { Length: > 0 } name
            ? name
            : string.IsNullOrWhiteSpace(ConnectionStringTextBox.Text)
                ? "No connection set"
                : "Custom connection";
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
            ConnectionStatusIndicator.Fill = (Brush)FindResource("AccentBrush");
            if (_selectedConnectionProfileId is { } profileId)
            {
                await _connectionProfileStoreService.TouchLastUsedAsync(profileId, CancellationToken.None);
            }

            await LoadSchemaMetadataAsync(connectionString);
        }
        else
        {
            SetStatus($"Connection failed: {result.ErrorMessage}");
            ConnectionStatusIndicator.Fill = (Brush)FindResource("DangerBrush");
            _toastService.Show($"Connection failed: {result.ErrorMessage}", ToastKind.Error);
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

        var parameterNames = QueryBatchSplitter.Split(sqlToExecute)
            .SelectMany(QueryOutputModeParser.ExtractParameterNames)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();
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

    private void OnTemplateActivated(string name, string sql)
    {
        SetQueryEditorText(sql);
        TrackRecentSqlFragments(sql);
        OutputTabControl.SelectedIndex = OutputSqlEditorTabIndex;
    }

    private bool ConfirmDeleteTemplate(string name)
    {
        var confirmation = MessageBox.Show(
            $"Delete template '{name}'?",
            "Delete Template",
            MessageBoxButton.YesNo,
            MessageBoxImage.Warning);

        return confirmation == MessageBoxResult.Yes;
    }

    private void OnSchemaTableSelected(TableSchemaInfo table, List<ColumnSchemaInfo> columns)
    {
        _selectedTable = table;
        _selectedColumns = columns;
        UpdateEditQueryTextFromInputs();
        ShowSelectedTableColumnsInEditRowsGrid();
        OutputTabControl.SelectedIndex = OutputEditRowsTabIndex;
        SubstituteTableNamePlaceholder(table);
    }

    private void OnSchemaTableCleared()
    {
        _selectedTable = null;
        _selectedColumns.Clear();
        UpdateEditQueryTextFromInputs();
        ExitEditModeAndClearEditableRows();
        EditRowsDataGrid.ItemsSource = null;
    }

    private void OnSchemaProcedureSelected(StoredProcedureSchemaInfo procedure, List<StoredProcedureParameterInfo> parameters)
    {
        _selectedStoredProcedure = procedure;
        _selectedProcedureParameters = parameters;
        OutputTabControl.SelectedIndex = OutputSchemaTabIndex;
    }

    private void OnSchemaProcedureCleared()
    {
        _selectedStoredProcedure = null;
        _selectedProcedureParameters.Clear();
    }

    private void OnSchemaMetadataLoaded(List<TableSchemaInfo> tables, List<StoredProcedureSchemaInfo> storedProcedures, List<ForeignKeySchemaInfo> foreignKeys)
    {
        _tables = tables;
        _storedProcedures = storedProcedures;
        _foreignKeys = foreignKeys;
        _selectedTable = null;
        _selectedStoredProcedure = null;
        _selectedColumns.Clear();
        _selectedProcedureParameters.Clear();
        ExitEditModeAndClearEditableRows();
        EditRowsDataGrid.ItemsSource = null;
    }

    private void OnSchemaScriptGenerated(string sql, string status)
    {
        SetQueryEditorText(sql);
        OutputTabControl.SelectedIndex = OutputSqlEditorTabIndex;
        SetStatus(status);
    }

    private void OnOpenInRunnerRequested(StoredProcedureSchemaInfo procedure, List<StoredProcedureParameterInfo> parameters)
    {
        _runnerParameterRows.Clear();

        foreach (var parameter in parameters.Where(x => !x.IsReturnValue))
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

        RunnerProcedureTextBlock.Text = $"Ready to execute {procedure.FullName}";
        OutputTabControl.SelectedIndex = OutputProcedureRunnerTabIndex;
        SetStatus("Loaded procedure parameters into runner.");
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
        var dialog = new DeleteColumnsSelectionWindow(availableColumns) { Owner = this };

        if (dialog.ShowDialog() != true)
        {
            selectedColumns = Array.Empty<string>();
            return false;
        }

        selectedColumns = dialog.SelectedColumns;
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


    private void SubstituteTableNamePlaceholder(TableSchemaInfo selectedTable)
    {
        var editors = new[] { _sqlEditor, _editRowsSqlEditor };
        foreach (var editor in editors)
        {
            if (editor?.Text is null)
            {
                continue;
            }

            var newText = TableNamePlaceholderSubstitution.TrySubstitute(editor.Text, selectedTable.SchemaName, selectedTable.TableName);
            if (newText is null)
            {
                continue;
            }

            var caretPos = editor.CaretIndex;
            editor.Text = newText;
            editor.CaretIndex = Math.Min(caretPos, editor.Text.Length);
        }
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
        await ViewModel.SchemaAssistant.LoadAsync(connectionString);
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
            _toastService.Show($"{operationName} failed: {result.ErrorMessage}", ToastKind.Error);
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
        StatusProgressBar.Visibility = isExecuting ? Visibility.Visible : Visibility.Collapsed;
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

    private async Task CopySchemaObjectNameToClipboardAsync(string text)
    {
        if (await TrySetClipboardTextAsync(text))
        {
            SetStatus($"Copied value to clipboard: {text}");
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
            // Covers both orderings: type/load the placeholder while a table is already
            // selected (this), or select a table after the placeholder is already there
            // (handled by OnSchemaTableSelected). SubstituteTableNamePlaceholder is a no-op
            // once the placeholder is gone, so the re-entrant TextChanged this triggers
            // terminates after one extra pass.
            if (_selectedTable is not null)
            {
                SubstituteTableNamePlaceholder(_selectedTable);
            }

            if (_isSettingQueryTextProgrammatically)
            {
                return;
            }

            _ = UpdateSqlSuggestionsForAsync(_sqlEditor!);
            return;
        }

        if (ReferenceEquals(sender, EditQueryTextBox))
        {
            _ = UpdateSqlSuggestionsForAsync(_editRowsSqlEditor!);
        }
    }

    /// <summary>
    /// Sets the SQL Editor's text without triggering the autocomplete suggestion popup -
    /// use this instead of "QueryTextBox.Text = ..." for any programmatic replacement
    /// (loading a template, generating a script from schema metadata, etc.), since a plain
    /// assignment fires the same TextChanged path as user typing and pops suggestions
    /// unrelated to what the user just did.
    /// </summary>
    private void SetQueryEditorText(string sql)
    {
        _isSettingQueryTextProgrammatically = true;
        try
        {
            QueryTextBox.Text = sql;
        }
        finally
        {
            _isSettingQueryTextProgrammatically = false;
        }

        HideSqlSuggestions();
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
        var dialog = new QueryParametersWindow(parameterNames) { Owner = this };

        if (dialog.ShowDialog() != true)
        {
            parameters = Array.Empty<QueryParameterValue>();
            return false;
        }

        parameters = dialog.Parameters;
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
}
