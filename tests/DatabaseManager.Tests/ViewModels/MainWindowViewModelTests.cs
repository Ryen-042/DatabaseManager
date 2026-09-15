using DatabaseManager.Wpf.Commands;
using DatabaseManager.Wpf.ViewModels;

namespace DatabaseManager.Tests.ViewModels;

public sealed class MainWindowViewModelTests
{
    [Fact]
    public void Defaults_AreDarkModeAndSchemaAssistantVisible()
    {
        var vm = new MainWindowViewModel(new AppCommandRegistry(), _ => { }, _ => { });

        Assert.True(vm.IsDarkMode);
        Assert.True(vm.IsSchemaAssistantVisible);
    }

    [Fact]
    public void SettingIsDarkMode_InvokesCallbackWithNewValue()
    {
        bool? observed = null;
        var vm = new MainWindowViewModel(new AppCommandRegistry(), value => observed = value, _ => { });

        vm.IsDarkMode = false;

        Assert.False(observed);
    }

    [Fact]
    public void SettingIsSchemaAssistantVisible_InvokesCallbackWithNewValue()
    {
        bool? observed = null;
        var vm = new MainWindowViewModel(new AppCommandRegistry(), _ => { }, value => observed = value);

        vm.IsSchemaAssistantVisible = false;

        Assert.False(observed);
    }

    [Fact]
    public void SettingSameValue_DoesNotInvokeCallback()
    {
        var invocationCount = 0;
        var vm = new MainWindowViewModel(new AppCommandRegistry(), _ => invocationCount++, _ => { });

        vm.IsDarkMode = true; // already the default value

        Assert.Equal(0, invocationCount);
    }

    [Fact]
    public void CommandRegistry_IsExposedAsGiven()
    {
        var registry = new AppCommandRegistry();
        var vm = new MainWindowViewModel(registry, _ => { }, _ => { });

        Assert.Same(registry, vm.CommandRegistry);
    }
}
