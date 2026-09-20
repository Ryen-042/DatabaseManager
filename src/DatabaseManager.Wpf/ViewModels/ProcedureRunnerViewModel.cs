using System.Collections.ObjectModel;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using DatabaseManager.Core.Models;
using DatabaseManager.Core.Models.Schema;
using DatabaseManager.Core.Services.Schema;

namespace DatabaseManager.Wpf.ViewModels;

/// <summary>
/// Procedure Runner tab: the parameter grid and the Execute command. "Open In Runner" (Schema
/// Assistant) populates this via LoadParameters; MainWindow still owns switching to this tab
/// and status text, reported through the same callback-injection pattern as the rest of the
/// extracted ViewModels.
///
/// Executes whichever procedure's parameters are currently loaded here (tracked internally),
/// not whatever happens to be selected in the Schema Assistant at click time - those can
/// diverge if the user browses to a different procedure after opening this one in the runner.
/// </summary>
public sealed partial class ProcedureRunnerViewModel : ObservableObject
{
    private readonly IStoredProcedureExecutionService _executionService;
    private readonly Func<string?> _getConnectionString;
    private readonly Func<bool> _getFullOutputCheckboxState;
    private readonly Action<bool> _setFullOutputMode;
    private readonly Func<int> _getTimeoutSeconds;
    private readonly Action<string> _setStatus;
    private readonly Action<string, QueryExecutionResult> _onResult;
    private readonly Action<bool> _onBusyChanged;

    private StoredProcedureSchemaInfo? _procedure;

    public ProcedureRunnerViewModel(
        IStoredProcedureExecutionService executionService,
        Func<string?> getConnectionString,
        Func<bool> getFullOutputCheckboxState,
        Action<bool> setFullOutputMode,
        Func<int> getTimeoutSeconds,
        Action<string> setStatus,
        Action<string, QueryExecutionResult> onResult,
        Action<bool> onBusyChanged)
    {
        _executionService = executionService;
        _getConnectionString = getConnectionString;
        _getFullOutputCheckboxState = getFullOutputCheckboxState;
        _setFullOutputMode = setFullOutputMode;
        _getTimeoutSeconds = getTimeoutSeconds;
        _setStatus = setStatus;
        _onResult = onResult;
        _onBusyChanged = onBusyChanged;
    }

    public ObservableCollection<ProcedureParameterEditorRow> Parameters { get; } = new();

    [ObservableProperty]
    private string _procedureSummary = "No procedure selected.";

    [ObservableProperty]
    private bool _isBusy;

    public void LoadParameters(StoredProcedureSchemaInfo procedure, IReadOnlyList<StoredProcedureParameterInfo> parameters)
    {
        _procedure = procedure;
        Parameters.Clear();

        foreach (var parameter in parameters.Where(x => !x.IsReturnValue))
        {
            Parameters.Add(new ProcedureParameterEditorRow
            {
                ParameterName = parameter.ParameterName,
                DataType = parameter.DataType,
                Value = string.Empty,
                SendAsNull = false,
                IsOutput = parameter.IsOutput
            });
        }

        ProcedureSummary = $"Ready to execute {procedure.FullName}";
    }

    partial void OnIsBusyChanged(bool value)
    {
        ExecuteCommand.NotifyCanExecuteChanged();
        _onBusyChanged(value);
    }

    private bool CanExecute() => !IsBusy;

    [RelayCommand(CanExecute = nameof(CanExecute))]
    private async Task ExecuteAsync()
    {
        if (_procedure is null)
        {
            _setStatus("Select a stored procedure first.");
            return;
        }

        var connectionString = _getConnectionString();
        if (string.IsNullOrWhiteSpace(connectionString))
        {
            _setStatus("Connection string is required.");
            return;
        }

        var fullOutputEnabled = _getFullOutputCheckboxState();
        _setFullOutputMode(fullOutputEnabled);

        IsBusy = true;
        _setStatus(fullOutputEnabled ? "Executing stored procedure (full output mode)..." : "Executing stored procedure...");

        var parameters = Parameters
            .Select(x => ProcedureParameterMapper.ToExecutionParameter(x.ParameterName, x.Value, x.SendAsNull, x.IsOutput))
            .ToList();

        var result = await _executionService.ExecuteAsync(
            connectionString,
            _procedure.SchemaName,
            _procedure.ProcedureName,
            parameters,
            _getTimeoutSeconds(),
            CancellationToken.None);

        _onResult("Stored procedure", result);
        IsBusy = false;
    }
}
