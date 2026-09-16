using System.Windows;
using System.Windows.Input;
using DatabaseManager.Wpf.Commands;

namespace DatabaseManager.Wpf.Windows;

public partial class CommandPaletteWindow : Window
{
    private readonly IReadOnlyList<AppCommandDescriptor> _allCommands;
    private bool _isClosing;

    public CommandPaletteWindow(IReadOnlyList<AppCommandDescriptor> commands)
    {
        InitializeComponent();
        _allCommands = commands;
        Loaded += (_, _) => FilterTextBox.Focus();
        Refilter(string.Empty);
    }

    private void FilterTextBox_TextChanged(object sender, System.Windows.Controls.TextChangedEventArgs e)
    {
        Refilter(FilterTextBox.Text);
    }

    private void Refilter(string query)
    {
        var matches = _allCommands
            .Select(c => (Command: c, Score: FuzzyMatcher.Score(query, $"{c.Category} {c.DisplayName} {string.Join(' ', c.Keywords)}")))
            .Where(m => m.Score.HasValue)
            .OrderByDescending(m => m.Score)
            .ThenBy(m => m.Command.Category, StringComparer.OrdinalIgnoreCase)
            .ThenBy(m => m.Command.DisplayName, StringComparer.OrdinalIgnoreCase)
            .Select(m => new PaletteItem(m.Command))
            .ToList();

        ResultsListBox.ItemsSource = matches;
        if (matches.Count > 0)
        {
            ResultsListBox.SelectedIndex = 0;
        }
    }

    private void CommandPaletteWindow_PreviewKeyDown(object sender, KeyEventArgs e)
    {
        switch (e.Key)
        {
            case Key.Escape:
                RequestClose();
                e.Handled = true;
                break;
            case Key.Down:
                MoveSelection(1);
                e.Handled = true;
                break;
            case Key.Up:
                MoveSelection(-1);
                e.Handled = true;
                break;
            case Key.Enter:
                ExecuteSelected();
                e.Handled = true;
                break;
        }
    }

    private void MoveSelection(int delta)
    {
        if (ResultsListBox.Items.Count == 0)
        {
            return;
        }

        var next = ResultsListBox.SelectedIndex + delta;
        next = Math.Clamp(next, 0, ResultsListBox.Items.Count - 1);
        ResultsListBox.SelectedIndex = next;
        ResultsListBox.ScrollIntoView(ResultsListBox.SelectedItem);
    }

    private void ResultsListBox_MouseDoubleClick(object sender, System.Windows.Input.MouseButtonEventArgs e)
    {
        ExecuteSelected();
    }

    private void ExecuteSelected()
    {
        if (ResultsListBox.SelectedItem is PaletteItem item && item.Descriptor.Command.CanExecute(null))
        {
            RequestClose();
            item.Descriptor.Command.Execute(null);
        }
    }

    private void CommandPaletteWindow_Deactivated(object sender, EventArgs e)
    {
        RequestClose();
    }

    /// <summary>
    /// Closing this window (a modal ShowDialog) deactivates it as part of the normal close
    /// sequence, which re-enters CommandPaletteWindow_Deactivated and would call Close() a
    /// second time on a window that's already closing - that reentrant call is what caused
    /// the app to hang for a couple of seconds and then crash when dismissing the palette
    /// with Escape. Every path that wants to close this window goes through here instead of
    /// calling Close() directly, so it only ever happens once.
    /// </summary>
    private void RequestClose()
    {
        if (_isClosing)
        {
            return;
        }

        _isClosing = true;
        Close();
    }

    private sealed class PaletteItem
    {
        public PaletteItem(AppCommandDescriptor descriptor)
        {
            Descriptor = descriptor;
        }

        public AppCommandDescriptor Descriptor { get; }

        public string Category => Descriptor.Category;

        public string DisplayName => Descriptor.DisplayName;

        public string GestureDisplayString => KeyGestureFormatter.Format(Descriptor.Gesture);
    }
}
