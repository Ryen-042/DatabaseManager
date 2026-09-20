using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using DatabaseManager.Wpf.ViewModels;

namespace DatabaseManager.Wpf.Views;

public partial class TablesPanelView : UserControl
{
    public TablesPanelView() => InitializeComponent();

    private void TablesListBox_PreviewMouseLeftButtonUp(object sender, MouseButtonEventArgs e)
    {
        if (DataContext is not SchemaAssistantViewModel viewModel)
        {
            return;
        }

        if (FindAncestor<ListBoxItem>(e.OriginalSource as DependencyObject)?.DataContext is not TableItemViewModel clicked)
        {
            return;
        }

        if (ReferenceEquals(viewModel.SelectedTableItem, clicked))
        {
            viewModel.ReapplySelectedTable();
        }
    }

    private static T? FindAncestor<T>(DependencyObject? current) where T : DependencyObject
    {
        while (current is not null)
        {
            if (current is T match)
            {
                return match;
            }

            current = VisualTreeHelper.GetParent(current);
        }

        return null;
    }
}
