using System.Windows;
using DatabaseManager.Wpf.Commands;

namespace DatabaseManager.Wpf.Windows;

public partial class ShortcutsHelpWindow : Window
{
    public ShortcutsHelpWindow(IReadOnlyList<AppCommandDescriptor> commands)
    {
        InitializeComponent();

        var groups = commands
            .Where(c => c.Gesture is not null)
            .GroupBy(c => c.Category, StringComparer.OrdinalIgnoreCase)
            .OrderBy(g => g.Key, StringComparer.OrdinalIgnoreCase)
            .Select(g => new CategoryGroup(
                g.Key,
                g.OrderBy(c => c.DisplayName, StringComparer.OrdinalIgnoreCase)
                    .Select(c => new ShortcutItem(c.DisplayName, KeyGestureFormatter.Format(c.Gesture)))
                    .ToList()))
            .ToList();

        CategoriesItemsControl.ItemsSource = groups;
    }

    private sealed record CategoryGroup(string Category, IReadOnlyList<ShortcutItem> Items);

    private sealed record ShortcutItem(string DisplayName, string GestureDisplayString);
}
