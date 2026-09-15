using System.Windows;
using System.Windows.Controls;
using DatabaseManager.Wpf.Commands;

namespace DatabaseManager.Wpf.Controls;

/// <summary>
/// Top menu whose items are built entirely from the shared <see cref="ICommandRegistry"/> -
/// grouped by each command's <see cref="AppCommandDescriptor.Category"/> - so it can never
/// list a command the palette/shortcuts panel don't also know about, or vice versa.
/// </summary>
public partial class AppMenu : UserControl
{
    public static readonly DependencyProperty CommandRegistryProperty = DependencyProperty.Register(
        nameof(CommandRegistry),
        typeof(ICommandRegistry),
        typeof(AppMenu),
        new PropertyMetadata(null, OnCommandRegistryChanged));

    public AppMenu()
    {
        InitializeComponent();
    }

    public ICommandRegistry? CommandRegistry
    {
        get => (ICommandRegistry?)GetValue(CommandRegistryProperty);
        set => SetValue(CommandRegistryProperty, value);
    }

    private static void OnCommandRegistryChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
    {
        var menu = (AppMenu)d;

        if (e.OldValue is ICommandRegistry oldRegistry)
        {
            oldRegistry.CommandsChanged -= menu.Registry_CommandsChanged;
        }

        if (e.NewValue is ICommandRegistry newRegistry)
        {
            newRegistry.CommandsChanged += menu.Registry_CommandsChanged;
        }

        menu.Rebuild();
    }

    private void Registry_CommandsChanged(object? sender, EventArgs e) => Rebuild();

    private void Rebuild()
    {
        RootMenu.Items.Clear();

        if (CommandRegistry is null)
        {
            return;
        }

        foreach (var group in CommandRegistry.Commands.GroupBy(c => c.Category, StringComparer.OrdinalIgnoreCase))
        {
            var categoryMenuItem = new MenuItem { Header = group.Key };

            foreach (var descriptor in group)
            {
                var itemMenuItem = new MenuItem
                {
                    Header = descriptor.DisplayName,
                    Command = descriptor.Command,
                    InputGestureText = KeyGestureFormatter.Format(descriptor.Gesture)
                };

                categoryMenuItem.Items.Add(itemMenuItem);
            }

            RootMenu.Items.Add(categoryMenuItem);
        }
    }
}
