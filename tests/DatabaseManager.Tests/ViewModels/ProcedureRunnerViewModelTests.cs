using DatabaseManager.Core.Models;
using DatabaseManager.Core.Models.Schema;
using DatabaseManager.Core.Services.Schema;
using DatabaseManager.Wpf.ViewModels;

namespace DatabaseManager.Tests.ViewModels;

public sealed class ProcedureRunnerViewModelTests
{
    private sealed class FakeStoredProcedureExecutionService : IStoredProcedureExecutionService
    {
        public QueryExecutionResult Result { get; set; } = new() { IsSuccess = true };
        public string? LastSchemaName { get; private set; }
        public string? LastProcedureName { get; private set; }
        public IReadOnlyList<StoredProcedureExecutionParameter>? LastParameters { get; private set; }
        public TaskCompletionSource<bool>? Gate { get; set; }

        public async Task<QueryExecutionResult> ExecuteAsync(string connectionString, string schemaName, string procedureName, IReadOnlyList<StoredProcedureExecutionParameter> parameters, int commandTimeoutSeconds, CancellationToken cancellationToken)
        {
            LastSchemaName = schemaName;
            LastProcedureName = procedureName;
            LastParameters = parameters;
            if (Gate is not null)
            {
                await Gate.Task;
            }

            return Result;
        }
    }

    private sealed class TestState
    {
        public List<string> Statuses { get; } = new();
        public List<(string Operation, QueryExecutionResult Result)> Results { get; } = new();
        public List<bool> BusyChanges { get; } = new();
        public bool? LastFullOutputMode { get; set; }
    }

    private static (ProcedureRunnerViewModel ViewModel, FakeStoredProcedureExecutionService Service, TestState State) Create(
        string? connectionString = "Server=test;",
        bool fullOutputCheckboxState = false)
    {
        var service = new FakeStoredProcedureExecutionService();
        var state = new TestState();

        var vm = new ProcedureRunnerViewModel(
            service,
            getConnectionString: () => connectionString,
            getFullOutputCheckboxState: () => fullOutputCheckboxState,
            setFullOutputMode: value => state.LastFullOutputMode = value,
            getTimeoutSeconds: () => 30,
            setStatus: state.Statuses.Add,
            onResult: (operation, result) => state.Results.Add((operation, result)),
            onBusyChanged: state.BusyChanges.Add);

        return (vm, service, state);
    }

    private static StoredProcedureSchemaInfo CreateProcedure(string name = "GetWidgets") => new() { SchemaName = "dbo", ProcedureName = name };

    private static List<StoredProcedureParameterInfo> CreateParameters() =>
    [
        new StoredProcedureParameterInfo { ParameterName = "@id", DataType = "int", OrdinalPosition = 1 },
        new StoredProcedureParameterInfo { ParameterName = "@out", DataType = "int", OrdinalPosition = 2, IsOutput = true },
        new StoredProcedureParameterInfo { ParameterName = "@RETURN_VALUE", DataType = "int", OrdinalPosition = 0 }
    ];

    [Fact]
    public void LoadParameters_ExcludesReturnValueAndSetsSummary()
    {
        var (vm, _, _) = Create();
        var procedure = CreateProcedure();

        vm.LoadParameters(procedure, CreateParameters());

        Assert.Equal(2, vm.Parameters.Count);
        Assert.DoesNotContain(vm.Parameters, p => p.ParameterName == "@RETURN_VALUE");
        Assert.Equal($"Ready to execute {procedure.FullName}", vm.ProcedureSummary);
    }

    [Fact]
    public void LoadParameters_ResetsValueAndSendAsNullForEachRow()
    {
        var (vm, _, _) = Create();

        vm.LoadParameters(CreateProcedure(), CreateParameters());

        Assert.All(vm.Parameters, p => Assert.Equal(string.Empty, p.Value));
        Assert.All(vm.Parameters, p => Assert.False(p.SendAsNull));
        Assert.True(vm.Parameters.Single(p => p.ParameterName == "@out").IsOutput);
    }

    [Fact]
    public async Task ExecuteAsync_NoProcedureLoaded_SetsStatusAndDoesNotCallService()
    {
        var (vm, service, state) = Create();

        await vm.ExecuteCommand.ExecuteAsync(null);

        Assert.Contains("Select a stored procedure first.", state.Statuses);
        Assert.Null(service.LastProcedureName);
    }

    [Fact]
    public async Task ExecuteAsync_MissingConnectionString_SetsStatusAndDoesNotCallService()
    {
        var (vm, service, state) = Create(connectionString: "");
        vm.LoadParameters(CreateProcedure(), CreateParameters());

        await vm.ExecuteCommand.ExecuteAsync(null);

        Assert.Contains("Connection string is required.", state.Statuses);
        Assert.Null(service.LastProcedureName);
    }

    [Fact]
    public async Task ExecuteAsync_PassesMappedParametersForTheLoadedProcedure()
    {
        var (vm, service, state) = Create();
        var procedure = CreateProcedure();
        vm.LoadParameters(procedure, CreateParameters());
        vm.Parameters.Single(p => p.ParameterName == "@id").Value = "42";

        await vm.ExecuteCommand.ExecuteAsync(null);

        Assert.Equal("dbo", service.LastSchemaName);
        Assert.Equal("GetWidgets", service.LastProcedureName);
        Assert.NotNull(service.LastParameters);
        var idParam = service.LastParameters!.Single(p => p.Name == "@id");
        Assert.Equal("42", idParam.Value);
        Assert.Single(state.Results);
        Assert.Equal("Stored procedure", state.Results[0].Operation);
    }

    [Fact]
    public async Task ExecuteAsync_TogglesIsBusyAroundExecution()
    {
        var (vm, service, state) = Create();
        vm.LoadParameters(CreateProcedure(), CreateParameters());
        service.Gate = new TaskCompletionSource<bool>();

        var executeTask = vm.ExecuteCommand.ExecuteAsync(null);
        Assert.True(vm.IsBusy);

        service.Gate.SetResult(true);
        await executeTask;

        Assert.False(vm.IsBusy);
        Assert.Equal(new[] { true, false }, state.BusyChanges);
    }

    [Fact]
    public async Task ExecuteAsync_FullOutputCheckboxState_IsReportedViaCallback()
    {
        var (vm, _, state) = Create(fullOutputCheckboxState: true);
        vm.LoadParameters(CreateProcedure(), CreateParameters());

        await vm.ExecuteCommand.ExecuteAsync(null);

        Assert.True(state.LastFullOutputMode);
    }

    [Fact]
    public async Task ExecuteAsync_ExecutesTheProcedureWhoseParametersAreLoaded_NotWhateverWasLoadedBefore()
    {
        var (vm, service, _) = Create();
        vm.LoadParameters(CreateProcedure("First"), CreateParameters());
        vm.LoadParameters(CreateProcedure("Second"), CreateParameters());

        await vm.ExecuteCommand.ExecuteAsync(null);

        Assert.Equal("Second", service.LastProcedureName);
    }
}
