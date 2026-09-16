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

    [Fact]
    public void Defaults_AreDarkModeAndSchemaAssistantVisible()
    {
        var vm = new MainWindowViewModel(new AppCommandRegistry(), _ => { }, _ => { }, CreateTemplatesPanel(), CreateSchemaAssistant());

        Assert.True(vm.IsDarkMode);
        Assert.True(vm.IsSchemaAssistantVisible);
    }

    [Fact]
    public void SettingIsDarkMode_InvokesCallbackWithNewValue()
    {
        bool? observed = null;
        var vm = new MainWindowViewModel(new AppCommandRegistry(), value => observed = value, _ => { }, CreateTemplatesPanel(), CreateSchemaAssistant());

        vm.IsDarkMode = false;

        Assert.False(observed);
    }

    [Fact]
    public void SettingIsSchemaAssistantVisible_InvokesCallbackWithNewValue()
    {
        bool? observed = null;
        var vm = new MainWindowViewModel(new AppCommandRegistry(), _ => { }, value => observed = value, CreateTemplatesPanel(), CreateSchemaAssistant());

        vm.IsSchemaAssistantVisible = false;

        Assert.False(observed);
    }

    [Fact]
    public void SettingSameValue_DoesNotInvokeCallback()
    {
        var invocationCount = 0;
        var vm = new MainWindowViewModel(new AppCommandRegistry(), _ => invocationCount++, _ => { }, CreateTemplatesPanel(), CreateSchemaAssistant());

        vm.IsDarkMode = true; // already the default value

        Assert.Equal(0, invocationCount);
    }

    [Fact]
    public void CommandRegistry_IsExposedAsGiven()
    {
        var registry = new AppCommandRegistry();
        var vm = new MainWindowViewModel(registry, _ => { }, _ => { }, CreateTemplatesPanel(), CreateSchemaAssistant());

        Assert.Same(registry, vm.CommandRegistry);
    }

    [Fact]
    public void TemplatesPanel_IsExposedAsGiven()
    {
        var templatesPanel = CreateTemplatesPanel();
        var vm = new MainWindowViewModel(new AppCommandRegistry(), _ => { }, _ => { }, templatesPanel, CreateSchemaAssistant());

        Assert.Same(templatesPanel, vm.TemplatesPanel);
    }

    [Fact]
    public void SchemaAssistant_IsExposedAsGiven()
    {
        var schemaAssistant = CreateSchemaAssistant();
        var vm = new MainWindowViewModel(new AppCommandRegistry(), _ => { }, _ => { }, CreateTemplatesPanel(), schemaAssistant);

        Assert.Same(schemaAssistant, vm.SchemaAssistant);
    }
}
