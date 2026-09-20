using DatabaseManager.Core.Models;
using DatabaseManager.Core.Services;
using DatabaseManager.Core.Services.Schema;
using DatabaseManager.Wpf.Commands;
using DatabaseManager.Wpf.SqlSuggestions;
using DatabaseManager.Wpf.ViewModels;

namespace DatabaseManager.Tests.ViewModels;

public sealed class MainWindowViewModelTests
{
    private static TemplatesPanelViewModel CreateTemplatesPanel() => new(
        new TemplateStoreService(Path.Combine(Path.GetTempPath(), $"dbm-unused-{Guid.NewGuid():N}.json")),
        getCurrentSqlText: () => string.Empty,
        onTemplateActivated: (_, _) => { },
        confirmDelete: _ => true,
        setStatus: _ => { });

    private static SchemaAssistantViewModel CreateSchemaAssistant() => new(
        new SqlServerSchemaService(),
        new SqlQueryAssistantService(),
        new SqlCompletionCatalogService(),
        getConnectionString: () => string.Empty,
        onTableSelected: (_, _) => { },
        onTableCleared: () => { },
        onProcedureSelected: (_, _) => { },
        onProcedureCleared: () => { },
        onSchemaMetadataLoaded: (_, _, _) => { },
        onScriptGenerated: (_, _) => { },
        onOpenInRunnerRequested: (_, _) => { },
        onCopyRequested: _ => Task.CompletedTask,
        setStatus: _ => { });

    private static QueryDocumentViewModel CreateQueryDocument() => new(
        new NullDatabaseQueryService(),
        getConnectionString: () => string.Empty,
        getSqlText: () => string.Empty,
        getFullOutputCheckboxState: () => false,
        setFullOutputMode: _ => { },
        getTimeoutSeconds: () => 30,
        promptForParameters: _ => (true, Array.Empty<QueryParameterValue>()),
        trackRecentSqlFragments: _ => { },
        setStatus: _ => { },
        onResult: (_, _) => { },
        onBusyChanged: _ => { });

    private sealed class NullDatabaseQueryService : IDatabaseQueryService
    {
        public Task<QueryExecutionResult> ExecuteAsync(string connectionString, string sql, int commandTimeoutSeconds, CancellationToken cancellationToken)
            => Task.FromResult(new QueryExecutionResult { IsSuccess = true });

        public Task<QueryExecutionResult> ExecuteAsync(string connectionString, string sql, IReadOnlyList<QueryParameterValue> parameters, int commandTimeoutSeconds, CancellationToken cancellationToken)
            => Task.FromResult(new QueryExecutionResult { IsSuccess = true });
    }

    [Fact]
    public void Defaults_AreDarkModeAndSchemaAssistantVisible()
    {
        var vm = new MainWindowViewModel(new AppCommandRegistry(), _ => { }, _ => { }, CreateTemplatesPanel(), CreateSchemaAssistant(), CreateQueryDocument());

        Assert.True(vm.IsDarkMode);
        Assert.True(vm.IsSchemaAssistantVisible);
    }

    [Fact]
    public void SettingIsDarkMode_InvokesCallbackWithNewValue()
    {
        bool? observed = null;
        var vm = new MainWindowViewModel(new AppCommandRegistry(), value => observed = value, _ => { }, CreateTemplatesPanel(), CreateSchemaAssistant(), CreateQueryDocument());

        vm.IsDarkMode = false;

        Assert.False(observed);
    }

    [Fact]
    public void SettingIsSchemaAssistantVisible_InvokesCallbackWithNewValue()
    {
        bool? observed = null;
        var vm = new MainWindowViewModel(new AppCommandRegistry(), _ => { }, value => observed = value, CreateTemplatesPanel(), CreateSchemaAssistant(), CreateQueryDocument());

        vm.IsSchemaAssistantVisible = false;

        Assert.False(observed);
    }

    [Fact]
    public void SettingSameValue_DoesNotInvokeCallback()
    {
        var invocationCount = 0;
        var vm = new MainWindowViewModel(new AppCommandRegistry(), _ => invocationCount++, _ => { }, CreateTemplatesPanel(), CreateSchemaAssistant(), CreateQueryDocument());

        vm.IsDarkMode = true; // already the default value

        Assert.Equal(0, invocationCount);
    }

    [Fact]
    public void CommandRegistry_IsExposedAsGiven()
    {
        var registry = new AppCommandRegistry();
        var vm = new MainWindowViewModel(registry, _ => { }, _ => { }, CreateTemplatesPanel(), CreateSchemaAssistant(), CreateQueryDocument());

        Assert.Same(registry, vm.CommandRegistry);
    }

    [Fact]
    public void TemplatesPanel_IsExposedAsGiven()
    {
        var templatesPanel = CreateTemplatesPanel();
        var vm = new MainWindowViewModel(new AppCommandRegistry(), _ => { }, _ => { }, templatesPanel, CreateSchemaAssistant(), CreateQueryDocument());

        Assert.Same(templatesPanel, vm.TemplatesPanel);
    }

    [Fact]
    public void SchemaAssistant_IsExposedAsGiven()
    {
        var schemaAssistant = CreateSchemaAssistant();
        var vm = new MainWindowViewModel(new AppCommandRegistry(), _ => { }, _ => { }, CreateTemplatesPanel(), schemaAssistant, CreateQueryDocument());

        Assert.Same(schemaAssistant, vm.SchemaAssistant);
    }

    [Fact]
    public void QueryDocument_IsExposedAsGiven()
    {
        var queryDocument = CreateQueryDocument();
        var vm = new MainWindowViewModel(new AppCommandRegistry(), _ => { }, _ => { }, CreateTemplatesPanel(), CreateSchemaAssistant(), queryDocument);

        Assert.Same(queryDocument, vm.QueryDocument);
    }
}
