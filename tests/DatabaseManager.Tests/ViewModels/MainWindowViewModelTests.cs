using DatabaseManager.Core.Services;
using DatabaseManager.Wpf.Commands;
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

    [Fact]
    public void Defaults_AreDarkModeAndSchemaAssistantVisible()
    {
        var vm = new MainWindowViewModel(new AppCommandRegistry(), _ => { }, _ => { }, CreateTemplatesPanel());

        Assert.True(vm.IsDarkMode);
        Assert.True(vm.IsSchemaAssistantVisible);
    }

    [Fact]
    public void SettingIsDarkMode_InvokesCallbackWithNewValue()
    {
        bool? observed = null;
        var vm = new MainWindowViewModel(new AppCommandRegistry(), value => observed = value, _ => { }, CreateTemplatesPanel());

        vm.IsDarkMode = false;

        Assert.False(observed);
    }

    [Fact]
    public void SettingIsSchemaAssistantVisible_InvokesCallbackWithNewValue()
    {
        bool? observed = null;
        var vm = new MainWindowViewModel(new AppCommandRegistry(), _ => { }, value => observed = value, CreateTemplatesPanel());

        vm.IsSchemaAssistantVisible = false;

        Assert.False(observed);
    }

    [Fact]
    public void SettingSameValue_DoesNotInvokeCallback()
    {
        var invocationCount = 0;
        var vm = new MainWindowViewModel(new AppCommandRegistry(), _ => invocationCount++, _ => { }, CreateTemplatesPanel());

        vm.IsDarkMode = true; // already the default value

        Assert.Equal(0, invocationCount);
    }

    [Fact]
    public void CommandRegistry_IsExposedAsGiven()
    {
        var registry = new AppCommandRegistry();
        var vm = new MainWindowViewModel(registry, _ => { }, _ => { }, CreateTemplatesPanel());

        Assert.Same(registry, vm.CommandRegistry);
    }

    [Fact]
    public void TemplatesPanel_IsExposedAsGiven()
    {
        var templatesPanel = CreateTemplatesPanel();
        var vm = new MainWindowViewModel(new AppCommandRegistry(), _ => { }, _ => { }, templatesPanel);

        Assert.Same(templatesPanel, vm.TemplatesPanel);
    }
}
