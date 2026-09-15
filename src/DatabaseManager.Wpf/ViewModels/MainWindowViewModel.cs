using CommunityToolkit.Mvvm.ComponentModel;
using DatabaseManager.Wpf.Commands;

namespace DatabaseManager.Wpf.ViewModels;

/// <summary>
/// Composition root for MainWindow's view-model state. Most of the app is still driven by
/// MainWindow.xaml.cs directly (see CLAUDE.md) - this class owns only the state that has been
/// migrated so far (theme, Object Explorer visibility) plus the shared command registry.
/// State changes are reported via callbacks the Window supplies, rather than this class
/// reaching into named XAML elements itself.
/// </summary>
public sealed partial class MainWindowViewModel : ObservableObject
{
    private readonly Action<bool> _onDarkModeChanged;
    private readonly Action<bool> _onSchemaAssistantVisibleChanged;

    public MainWindowViewModel(
        ICommandRegistry commandRegistry,
        Action<bool> onDarkModeChanged,
        Action<bool> onSchemaAssistantVisibleChanged)
    {
        CommandRegistry = commandRegistry;
        _onDarkModeChanged = onDarkModeChanged;
        _onSchemaAssistantVisibleChanged = onSchemaAssistantVisibleChanged;
    }

    public ICommandRegistry CommandRegistry { get; }

    [ObservableProperty]
    private bool _isDarkMode = true;

    [ObservableProperty]
    private bool _isSchemaAssistantVisible = true;

    partial void OnIsDarkModeChanged(bool value) => _onDarkModeChanged(value);

    partial void OnIsSchemaAssistantVisibleChanged(bool value) => _onSchemaAssistantVisibleChanged(value);
}
