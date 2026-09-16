using DatabaseManager.Core.Models.Schema;
using DatabaseManager.Core.Services.Schema;
using DatabaseManager.Wpf.SqlSuggestions;
using DatabaseManager.Wpf.ViewModels;

namespace DatabaseManager.Tests.ViewModels;

public sealed class SchemaAssistantViewModelTests
{
    private sealed class FakeDatabaseSchemaService : IDatabaseSchemaService
    {
        public List<TableSchemaInfo> Tables { get; set; } = new();
        public List<StoredProcedureSchemaInfo> StoredProcedures { get; set; } = new();
        public List<ForeignKeySchemaInfo> ForeignKeys { get; set; } = new();
        public List<ColumnSchemaInfo> Columns { get; set; } = new();
        public List<StoredProcedureParameterInfo> Parameters { get; set; } = new();
        public string? ProcedureDefinition { get; set; } = "CREATE PROCEDURE dbo.Foo AS SELECT 1;";
        public Exception? ThrowOnGetTables { get; set; }
        public Exception? ThrowOnGetColumns { get; set; }

        public Task<IReadOnlyList<TableSchemaInfo>> GetTablesAsync(string connectionString, CancellationToken cancellationToken)
            => ThrowOnGetTables is null
                ? Task.FromResult<IReadOnlyList<TableSchemaInfo>>(Tables)
                : Task.FromException<IReadOnlyList<TableSchemaInfo>>(ThrowOnGetTables);

        public Task<IReadOnlyList<ForeignKeySchemaInfo>> GetForeignKeysAsync(string connectionString, CancellationToken cancellationToken)
            => Task.FromResult<IReadOnlyList<ForeignKeySchemaInfo>>(ForeignKeys);

        public Task<IReadOnlyList<ColumnSchemaInfo>> GetColumnsAsync(string connectionString, string schemaName, string tableName, CancellationToken cancellationToken)
            => ThrowOnGetColumns is null
                ? Task.FromResult<IReadOnlyList<ColumnSchemaInfo>>(Columns)
                : Task.FromException<IReadOnlyList<ColumnSchemaInfo>>(ThrowOnGetColumns);

        public Task<IReadOnlyList<StoredProcedureSchemaInfo>> GetStoredProceduresAsync(string connectionString, CancellationToken cancellationToken)
            => Task.FromResult<IReadOnlyList<StoredProcedureSchemaInfo>>(StoredProcedures);

        public Task<IReadOnlyList<StoredProcedureParameterInfo>> GetStoredProcedureParametersAsync(string connectionString, string schemaName, string procedureName, CancellationToken cancellationToken)
            => Task.FromResult<IReadOnlyList<StoredProcedureParameterInfo>>(Parameters);

        public Task<string?> GetStoredProcedureDefinitionAsync(string connectionString, string schemaName, string procedureName, CancellationToken cancellationToken)
            => Task.FromResult(ProcedureDefinition);
    }

    private sealed class TestState
    {
        public List<(TableSchemaInfo Table, List<ColumnSchemaInfo> Columns)> TableSelected { get; } = new();
        public int TableClearedCount { get; set; }
        public List<(StoredProcedureSchemaInfo Procedure, List<StoredProcedureParameterInfo> Parameters)> ProcedureSelected { get; } = new();
        public int ProcedureClearedCount { get; set; }
        public List<(List<TableSchemaInfo> Tables, List<StoredProcedureSchemaInfo> Procedures, List<ForeignKeySchemaInfo> ForeignKeys)> MetadataLoaded { get; } = new();
        public List<(string Sql, string Status)> ScriptsGenerated { get; } = new();
        public List<(StoredProcedureSchemaInfo Procedure, List<StoredProcedureParameterInfo> Parameters)> OpenInRunnerRequests { get; } = new();
        public List<string> CopyRequests { get; } = new();
        public List<string> Statuses { get; } = new();
    }

    private static (SchemaAssistantViewModel ViewModel, FakeDatabaseSchemaService Schema, TestState State) Create(string? connectionString = "Server=test;")
    {
        var schema = new FakeDatabaseSchemaService();
        var state = new TestState();

        var vm = new SchemaAssistantViewModel(
            schema,
            new SqlQueryAssistantService(),
            new SqlCompletionCatalogService(),
            getConnectionString: () => connectionString,
            onTableSelected: (table, columns) => state.TableSelected.Add((table, columns)),
            onTableCleared: () => state.TableClearedCount++,
            onProcedureSelected: (procedure, parameters) => state.ProcedureSelected.Add((procedure, parameters)),
            onProcedureCleared: () => state.ProcedureClearedCount++,
            onSchemaMetadataLoaded: (tables, procedures, foreignKeys) => state.MetadataLoaded.Add((tables, procedures, foreignKeys)),
            onScriptGenerated: (sql, status) => state.ScriptsGenerated.Add((sql, status)),
            onOpenInRunnerRequested: (procedure, parameters) => state.OpenInRunnerRequests.Add((procedure, parameters)),
            onCopyRequested: text =>
            {
                state.CopyRequests.Add(text);
                return Task.CompletedTask;
            },
            setStatus: state.Statuses.Add);

        return (vm, schema, state);
    }

    private static TableSchemaInfo CreateTable(string name = "Widgets") => new() { SchemaName = "dbo", TableName = name };

    private static StoredProcedureSchemaInfo CreateProcedure(string name = "GetWidgets") => new() { SchemaName = "dbo", ProcedureName = name };

    private static ColumnSchemaInfo CreateColumn(string name, bool isPrimaryKey = false) => new()
    {
        SchemaName = "dbo",
        TableName = "Widgets",
        ColumnName = name,
        DataType = "int",
        IsPrimaryKey = isPrimaryKey
    };

    [Fact]
    public async Task LoadAsync_PopulatesFilteredListsAndStatus()
    {
        var (vm, schema, state) = Create();
        schema.Tables = [CreateTable("Widgets"), CreateTable("Gadgets")];
        schema.StoredProcedures = [CreateProcedure()];

        await vm.LoadAsync("Server=test;");

        Assert.Equal(2, vm.FilteredTables.Count);
        Assert.Single(vm.FilteredStoredProcedures);
        Assert.Single(state.MetadataLoaded);
        Assert.Contains("Schema loaded successfully", state.Statuses[^1]);
    }

    [Fact]
    public async Task LoadAsync_NoTablesOrProcedures_SetsEmptyStatus()
    {
        var (vm, _, state) = Create();

        await vm.LoadAsync("Server=test;");

        Assert.Contains("no tables or procedures were found", state.Statuses[^1]);
    }

    [Fact]
    public async Task LoadAsync_Failure_SetsErrorStatus()
    {
        var (vm, schema, state) = Create();
        schema.ThrowOnGetTables = new InvalidOperationException("boom");

        await vm.LoadAsync("Server=test;");

        Assert.Contains("Failed to load schema metadata: boom", state.Statuses[^1]);
    }

    [Fact]
    public async Task TableSearchText_FiltersTables()
    {
        var (vm, schema, _) = Create();
        schema.Tables = [CreateTable("Widgets"), CreateTable("Gadgets")];
        await vm.LoadAsync("Server=test;");

        vm.TableSearchText = "Gad";

        Assert.Single(vm.FilteredTables);
        Assert.Equal("[dbo].[Gadgets]", vm.FilteredTables[0].FullName);
    }

    [Fact]
    public async Task SelectingTable_LoadsColumnsAndInvokesCallback()
    {
        var (vm, schema, state) = Create();
        var table = CreateTable();
        schema.Tables = [table];
        schema.Columns = [CreateColumn("Id", isPrimaryKey: true)];
        await vm.LoadAsync("Server=test;");

        vm.SelectedTableItem = vm.FilteredTables[0];
        await Task.Delay(50); // selection change kicks off an async load

        Assert.Single(state.TableSelected);
        Assert.Equal(table.FullName, state.TableSelected[0].Table.FullName);
        Assert.Single(vm.SelectedColumns);
        Assert.Contains(table.FullName, vm.SchemaSummaryText);
    }

    [Fact]
    public async Task ClearingTableSelection_InvokesOnTableCleared()
    {
        var (vm, schema, state) = Create();
        schema.Tables = [CreateTable()];
        schema.Columns = [CreateColumn("Id", isPrimaryKey: true)];
        await vm.LoadAsync("Server=test;");
        vm.SelectedTableItem = vm.FilteredTables[0];
        await Task.Delay(50);

        vm.SelectedTableItem = null;

        Assert.Equal(1, state.TableClearedCount);
        Assert.Empty(vm.SelectedColumns);
    }

    [Fact]
    public async Task SelectingProcedure_LoadsParametersAndDefinitionAndInvokesCallback()
    {
        var (vm, schema, state) = Create();
        var procedure = CreateProcedure();
        schema.StoredProcedures = [procedure];
        schema.Parameters = [new StoredProcedureParameterInfo { ParameterName = "@id", DataType = "int", OrdinalPosition = 1 }];
        await vm.LoadAsync("Server=test;");

        vm.SelectedProcedureItem = vm.FilteredStoredProcedures[0];
        await Task.Delay(50);

        Assert.Single(state.ProcedureSelected);
        Assert.Single(vm.SelectedProcedureParameters);
        Assert.Equal(schema.ProcedureDefinition, vm.ProcedureSqlDefinitionText);
    }

    [Fact]
    public void GenerateSelect_NoTableSelected_SetsStatusAndDoesNotInvokeCallback()
    {
        var (vm, _, state) = Create();

        vm.GenerateSelectCommand.Execute(null);

        Assert.Empty(state.ScriptsGenerated);
        Assert.Contains("Select a table first.", state.Statuses);
    }

    [Fact]
    public async Task GenerateSelect_WithTableSelected_InvokesOnScriptGenerated()
    {
        var (vm, schema, state) = Create();
        schema.Tables = [CreateTable()];
        schema.Columns = [CreateColumn("Id", isPrimaryKey: true)];
        await vm.LoadAsync("Server=test;");
        vm.SelectedTableItem = vm.FilteredTables[0];
        await Task.Delay(50);

        vm.GenerateSelectCommand.Execute(null);

        Assert.Single(state.ScriptsGenerated);
        Assert.Contains("SELECT", state.ScriptsGenerated[0].Sql, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task GenerateUpdate_NoPrimaryKey_UsesTodoStatusMessage()
    {
        var (vm, schema, state) = Create();
        schema.Tables = [CreateTable()];
        schema.Columns = [CreateColumn("Name")];
        await vm.LoadAsync("Server=test;");
        vm.SelectedTableItem = vm.FilteredTables[0];
        await Task.Delay(50);

        vm.GenerateUpdateCommand.Execute(null);

        Assert.Single(state.ScriptsGenerated);
        Assert.Equal("Generated UPDATE query. No primary key detected, so WHERE clause needs manual fix.", state.ScriptsGenerated[0].Status);
    }

    [Fact]
    public async Task GenerateExec_WithProcedureSelected_InvokesOnScriptGenerated()
    {
        var (vm, schema, state) = Create();
        schema.StoredProcedures = [CreateProcedure()];
        await vm.LoadAsync("Server=test;");
        vm.SelectedProcedureItem = vm.FilteredStoredProcedures[0];
        await Task.Delay(50);

        vm.GenerateExecCommand.Execute(null);

        Assert.Single(state.ScriptsGenerated);
        Assert.Contains("EXEC", state.ScriptsGenerated[0].Sql, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task OpenInRunner_WithProcedureSelected_InvokesCallback()
    {
        var (vm, schema, state) = Create();
        schema.StoredProcedures = [CreateProcedure()];
        await vm.LoadAsync("Server=test;");
        vm.SelectedProcedureItem = vm.FilteredStoredProcedures[0];
        await Task.Delay(50);

        vm.OpenInRunnerCommand.Execute(null);

        Assert.Single(state.OpenInRunnerRequests);
    }

    [Fact]
    public void CopySelectedObjectName_NoSelection_SetsStatus()
    {
        var (vm, _, state) = Create();

        vm.CopySelectedObjectNameCommand.Execute(null);

        Assert.Contains("Select a table or stored procedure first.", state.Statuses);
        Assert.Empty(state.CopyRequests);
    }

    [Fact]
    public async Task CopySelectedObjectName_WithTableSelected_InvokesCallback()
    {
        var (vm, schema, state) = Create();
        var table = CreateTable();
        schema.Tables = [table];
        schema.Columns = [CreateColumn("Id", isPrimaryKey: true)];
        await vm.LoadAsync("Server=test;");
        vm.SelectedTableItem = vm.FilteredTables[0];
        await Task.Delay(50);

        await vm.CopySelectedObjectNameCommand.ExecuteAsync(null);

        Assert.Single(state.CopyRequests);
        Assert.Equal(table.FullName, state.CopyRequests[0]);
    }
}
