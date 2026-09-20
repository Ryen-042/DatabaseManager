using DatabaseManager.Core.Models;
using DatabaseManager.Core.Services;
using DatabaseManager.Wpf.ViewModels;

namespace DatabaseManager.Tests.ViewModels;

public sealed class QueryDocumentViewModelTests
{
    private sealed class FakeDatabaseQueryService : IDatabaseQueryService
    {
        public QueryExecutionResult Result { get; set; } = new() { IsSuccess = true, AffectedRows = 0 };
        public string? LastSql { get; private set; }
        public IReadOnlyList<QueryParameterValue>? LastParameters { get; private set; }
        public CancellationToken LastCancellationToken { get; private set; }
        public TaskCompletionSource<bool>? Gate { get; set; }

        public async Task<QueryExecutionResult> ExecuteAsync(string connectionString, string sql, int commandTimeoutSeconds, CancellationToken cancellationToken)
        {
            LastSql = sql;
            LastParameters = null;
            LastCancellationToken = cancellationToken;
            if (Gate is not null)
            {
                await Gate.Task;
            }

            // Matches SqlServerQueryService's real contract: cancellation is caught internally
            // and reported as a failed result, never thrown out of ExecuteAsync.
            return cancellationToken.IsCancellationRequested
                ? new QueryExecutionResult { IsSuccess = false, ErrorMessage = "Operation canceled." }
                : Result;
        }

        public async Task<QueryExecutionResult> ExecuteAsync(string connectionString, string sql, IReadOnlyList<QueryParameterValue> parameters, int commandTimeoutSeconds, CancellationToken cancellationToken)
        {
            LastSql = sql;
            LastParameters = parameters;
            LastCancellationToken = cancellationToken;
            if (Gate is not null)
            {
                await Gate.Task;
            }

            return cancellationToken.IsCancellationRequested
                ? new QueryExecutionResult { IsSuccess = false, ErrorMessage = "Operation canceled." }
                : Result;
        }
    }

    private sealed class TestState
    {
        public List<string> Statuses { get; } = new();
        public List<(string Operation, QueryExecutionResult Result)> Results { get; } = new();
        public List<bool> BusyChanges { get; } = new();
        public List<string> RecentFragments { get; } = new();
        public bool? LastFullOutputMode { get; set; }
    }

    private static (QueryDocumentViewModel ViewModel, FakeDatabaseQueryService Service, TestState State) Create(
        string sql = "SELECT 1;",
        string? connectionString = "Server=test;",
        bool fullOutputCheckboxState = false,
        Func<IReadOnlyList<string>, (bool Proceed, IReadOnlyList<QueryParameterValue> Parameters)>? promptForParameters = null)
    {
        var service = new FakeDatabaseQueryService();
        var state = new TestState();

        var vm = new QueryDocumentViewModel(
            service,
            getConnectionString: () => connectionString,
            getSqlText: () => sql,
            getFullOutputCheckboxState: () => fullOutputCheckboxState,
            setFullOutputMode: value => state.LastFullOutputMode = value,
            getTimeoutSeconds: () => 30,
            promptForParameters: promptForParameters ?? (_ => (true, Array.Empty<QueryParameterValue>())),
            trackRecentSqlFragments: state.RecentFragments.Add,
            setStatus: state.Statuses.Add,
            onResult: (operation, result) => state.Results.Add((operation, result)),
            onBusyChanged: state.BusyChanges.Add);

        return (vm, service, state);
    }

    [Fact]
    public async Task RunAsync_Success_InvokesResultCallbackAndClearsDirty()
    {
        var (vm, service, state) = Create();
        service.Result = new QueryExecutionResult { IsSuccess = true, AffectedRows = 1 };
        vm.MarkDirty();

        await vm.RunCommand.ExecuteAsync(null);

        Assert.False(vm.IsDirty);
        Assert.Single(state.Results);
        Assert.Equal("Query", state.Results[0].Operation);
        Assert.True(state.Results[0].Result.IsSuccess);
    }

    [Fact]
    public async Task RunAsync_Failure_KeepsDirtyTrue()
    {
        var (vm, service, _) = Create();
        service.Result = new QueryExecutionResult { IsSuccess = false, ErrorMessage = "boom" };
        vm.MarkDirty();

        await vm.RunCommand.ExecuteAsync(null);

        Assert.True(vm.IsDirty);
    }

    [Fact]
    public async Task RunAsync_TogglesIsBusyAroundExecution()
    {
        var (vm, service, state) = Create();
        service.Gate = new TaskCompletionSource<bool>();

        var runTask = vm.RunCommand.ExecuteAsync(null);
        Assert.True(vm.IsBusy);

        service.Gate.SetResult(true);
        await runTask;

        Assert.False(vm.IsBusy);
        Assert.Equal(new[] { true, false }, state.BusyChanges);
    }

    [Fact]
    public async Task RunAsync_MissingConnectionString_SetsStatusAndDoesNotExecute()
    {
        var (vm, service, state) = Create(connectionString: "");

        await vm.RunCommand.ExecuteAsync(null);

        Assert.Contains("Connection string is required.", state.Statuses);
        Assert.Null(service.LastSql);
    }

    [Fact]
    public async Task RunAsync_EmptySql_SetsStatusAndDoesNotExecute()
    {
        var (vm, service, state) = Create(sql: "   ");

        await vm.RunCommand.ExecuteAsync(null);

        Assert.Contains("SQL query is required.", state.Statuses);
        Assert.Null(service.LastSql);
    }

    [Fact]
    public async Task RunAsync_QueryWithParameters_PromptsAndPassesSuppliedValues()
    {
        var suppliedParameters = new List<QueryParameterValue> { new() { Name = "@id", Value = 5 } };
        var (vm, service, _) = Create(
            sql: "SELECT * FROM Users WHERE Id = @id;",
            promptForParameters: names =>
            {
                Assert.Contains("@id", names);
                return (true, suppliedParameters);
            });

        await vm.RunCommand.ExecuteAsync(null);

        Assert.Same(suppliedParameters, service.LastParameters);
    }

    [Fact]
    public async Task RunAsync_ParameterPromptCanceled_DoesNotExecute()
    {
        var (vm, service, state) = Create(
            sql: "SELECT * FROM Users WHERE Id = @id;",
            promptForParameters: _ => (false, Array.Empty<QueryParameterValue>()));

        await vm.RunCommand.ExecuteAsync(null);

        Assert.Contains("Query execution canceled.", state.Statuses);
        Assert.Null(service.LastSql);
    }

    [Fact]
    public async Task RunAsync_CreateProcedureWithParameters_DoesNotPrompt()
    {
        // Batch-aware: a CREATE PROC's own parameter declarations must not be treated as
        // query parameters to prompt for (this was the original bug this app started with).
        var (vm, service, state) = Create(
            sql: "CREATE PROC dbo.MyProc @Id INT AS BEGIN SELECT 1; END;",
            promptForParameters: _ => throw new InvalidOperationException("Should not prompt for a CREATE PROC statement."));

        await vm.RunCommand.ExecuteAsync(null);

        Assert.NotNull(service.LastSql);
        Assert.DoesNotContain("Query execution canceled.", state.Statuses);
    }

    [Fact]
    public async Task RunAsync_FullDirective_OverridesCheckboxAndReportsFullOutputMode()
    {
        var (vm, _, state) = Create(sql: "SELECT 1; -- full", fullOutputCheckboxState: false);

        await vm.RunCommand.ExecuteAsync(null);

        Assert.True(state.LastFullOutputMode);
    }

    [Fact]
    public void Cancel_RequestsCancellationOfInFlightRun()
    {
        var (vm, service, state) = Create();
        service.Gate = new TaskCompletionSource<bool>();

        var runTask = vm.RunCommand.ExecuteAsync(null);

        vm.CancelCommand.Execute(null);
        service.Gate.SetResult(true);

        Assert.Contains("Cancellation requested...", state.Statuses);
        Assert.True(service.LastCancellationToken.IsCancellationRequested);
    }
}
