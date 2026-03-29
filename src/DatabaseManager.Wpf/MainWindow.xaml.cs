using System.Data;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Interop;
using DatabaseManager.Core.Models;
using DatabaseManager.Core.Services;
using Microsoft.Win32;

namespace DatabaseManager.Wpf;

public partial class MainWindow : Window
{
    private const int DwmwaUseImmersiveDarkMode = 20;
    private const int DwmwaUseImmersiveDarkModeBefore20H1 = 19;

    private readonly IDatabaseQueryService _databaseQueryService = new SqlServerQueryService();
    private readonly ITemplateStoreService _templateStoreService = new TemplateStoreService();
    private readonly IExportService _exportService = new ExportService();

    private DataTable? _currentDataTable;
    private CancellationTokenSource? _executionCancellationTokenSource;

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
        var connectionString = ConnectionStringTextBox.Text.Trim();
        if (string.IsNullOrWhiteSpace(connectionString))
        {
            SetStatus("Connection string is required.");
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
        }
        else
        {
            SetStatus($"Connection test failed: {result.ErrorMessage}");
        }

        SetExecutionState(false);
    }

    private async void RunQueryButton_Click(object sender, RoutedEventArgs e)
    {
        var connectionString = ConnectionStringTextBox.Text.Trim();
        var sql = QueryTextBox.Text;

        if (string.IsNullOrWhiteSpace(connectionString))
        {
            SetStatus("Connection string is required.");
            return;
        }

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

        if (!result.IsSuccess)
        {
            SetStatus($"Query failed after {result.Duration.TotalSeconds:F2}s: {result.ErrorMessage}");
            SetExecutionState(false);
            return;
        }

        _currentDataTable = result.DataTable;
        ResultsDataGrid.ItemsSource = _currentDataTable?.DefaultView;
        ResultsSummaryTextBlock.Text = _currentDataTable is null
            ? $"Results ({result.AffectedRows} affected rows)"
            : $"Results ({_currentDataTable.Rows.Count} rows, {_currentDataTable.Columns.Count} columns)";

        SetStatus($"Query completed in {result.Duration.TotalSeconds:F2}s.");
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
}