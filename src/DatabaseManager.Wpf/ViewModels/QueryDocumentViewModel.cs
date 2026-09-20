using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using DatabaseManager.Core.Models;
using DatabaseManager.Core.Services;

namespace DatabaseManager.Wpf.ViewModels;

/// <summary>
/// Owns Run/Cancel execution state for the SQL Editor tab. The AvalonEdit control, the
/// results display, and the shared SQL suggestion popup (also used by Edit Rows) still live
/// in MainWindow - this class owns only the state that's been migrated so far (IsBusy,
/// IsDirty, and the Run/Cancel commands), reaching out via constructor-supplied callbacks for
/// anything outside its own concern, the same pattern SchemaAssistantViewModel established.
/// </summary>
public sealed partial class QueryDocumentViewModel : ObservableObject
{
    private readonly IDatabaseQueryService _queryService;
    private readonly Func<string?> _getConnectionString;
    private readonly Func<string> _getSqlText;
    private readonly Func<bool> _getFullOutputCheckboxState;
    private readonly Action<bool> _setFullOutputMode;
    private readonly Func<int> _getTimeoutSeconds;
    private readonly Func<IReadOnlyList<string>, (bool Proceed, IReadOnlyList<QueryParameterValue> Parameters)> _promptForParameters;
    private readonly Action<string> _trackRecentSqlFragments;
    private readonly Action<string> _setStatus;
    private readonly Action<string, QueryExecutionResult> _onResult;
    private readonly Action<bool> _onBusyChanged;

    private CancellationTokenSource? _executionCancellationTokenSource;

    public QueryDocumentViewModel(
        IDatabaseQueryService queryService,
        Func<string?> getConnectionString,
        Func<string> getSqlText,
        Func<bool> getFullOutputCheckboxState,
        Action<bool> setFullOutputMode,
        Func<int> getTimeoutSeconds,
        Func<IReadOnlyList<string>, (bool Proceed, IReadOnlyList<QueryParameterValue> Parameters)> promptForParameters,
        Action<string> trackRecentSqlFragments,
        Action<string> setStatus,
        Action<string, QueryExecutionResult> onResult,
        Action<bool> onBusyChanged)
    {
        _queryService = queryService;
        _getConnectionString = getConnectionString;
        _getSqlText = getSqlText;
        _getFullOutputCheckboxState = getFullOutputCheckboxState;
        _setFullOutputMode = setFullOutputMode;
        _getTimeoutSeconds = getTimeoutSeconds;
        _promptForParameters = promptForParameters;
        _trackRecentSqlFragments = trackRecentSqlFragments;
        _setStatus = setStatus;
        _onResult = onResult;
        _onBusyChanged = onBusyChanged;
    }

    [ObservableProperty]
    private bool _isBusy;

    /// <summary>
    /// True after any SQL text edit since the last successful run; reset to false immediately
    /// after a successful run. Not yet surfaced in the UI - this is the prerequisite Phase 7's
    /// per-document dirty tab indicator will read once multiple query documents exist.
    /// </summary>
    [ObservableProperty]
    private bool _isDirty;

    public void MarkDirty() => IsDirty = true;

    partial void OnIsBusyChanged(bool value)
    {
        RunCommand.NotifyCanExecuteChanged();
        CancelCommand.NotifyCanExecuteChanged();
        _onBusyChanged(value);
    }

    private bool CanRun() => !IsBusy;

    [RelayCommand(CanExecute = nameof(CanRun))]
    private async Task RunAsync()
    {
        var connectionString = _getConnectionString();
        if (string.IsNullOrWhiteSpace(connectionString))
        {
            _setStatus("Connection string is required.");
            return;
        }

        var sql = _getSqlText();
        if (string.IsNullOrWhiteSpace(sql))
        {
            _setStatus("SQL query is required.");
            return;
        }

        var outputMode = QueryOutputModeParser.Parse(sql);
        var sqlToExecute = outputMode.Sql;
        if (string.IsNullOrWhiteSpace(sqlToExecute))
        {
            _setStatus("SQL query is required.");
            return;
        }

        var fullOutputEnabled = outputMode.HasFullDirective || _getFullOutputCheckboxState();
        _setFullOutputMode(fullOutputEnabled);
        _trackRecentSqlFragments(sqlToExecute);

        var parameterNames = QueryBatchSplitter.Split(sqlToExecute)
            .SelectMany(QueryOutputModeParser.ExtractParameterNames)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();

        IReadOnlyList<QueryParameterValue> queryParameters = Array.Empty<QueryParameterValue>();
        if (parameterNames.Count > 0)
        {
            var (proceed, parameters) = _promptForParameters(parameterNames);
            if (!proceed)
            {
                _setStatus("Query execution canceled.");
                return;
            }

            queryParameters = parameters;
        }

        _executionCancellationTokenSource?.Dispose();
        _executionCancellationTokenSource = new CancellationTokenSource();

        IsBusy = true;
        _setStatus(fullOutputEnabled ? "Executing query (full output mode)..." : "Executing query...");

        var timeoutSeconds = _getTimeoutSeconds();
        var result = queryParameters.Count == 0
            ? await _queryService.ExecuteAsync(connectionString, sqlToExecute, timeoutSeconds, _executionCancellationTokenSource.Token)
            : await _queryService.ExecuteAsync(connectionString, sqlToExecute, queryParameters, timeoutSeconds, _executionCancellationTokenSource.Token);

        if (result.IsSuccess)
        {
            IsDirty = false;
        }

        _onResult("Query", result);
        IsBusy = false;
    }

    [RelayCommand(CanExecute = nameof(IsBusy))]
    private void Cancel()
    {
        _executionCancellationTokenSource?.Cancel();
        _setStatus("Cancellation requested...");
    }
}
