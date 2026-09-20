using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Threading;

namespace DatabaseManager.Wpf.Behaviors;

/// <summary>
/// A freshly-loaded DataGrid can start horizontally scrolled past its first column (WPF
/// restores the previous ItemsSource's scroll offset). This attached property resets it back
/// to the left on Loaded - previously duplicated as a private DataGrid_Loaded/
/// ResetDataGridHorizontalScroll/FindDescendant&lt;T&gt; trio in both MainWindow.xaml.cs and
/// SchemaDetailView.xaml.cs; this is the one copy now.
/// </summary>
public static class DataGridBehaviors
{
    public static readonly DependencyProperty ResetHorizontalScrollOnLoadProperty =
        DependencyProperty.RegisterAttached(
            "ResetHorizontalScrollOnLoad",
            typeof(bool),
            typeof(DataGridBehaviors),
            new PropertyMetadata(false, OnResetHorizontalScrollOnLoadChanged));

    public static void SetResetHorizontalScrollOnLoad(DependencyObject element, bool value)
        => element.SetValue(ResetHorizontalScrollOnLoadProperty, value);

    public static bool GetResetHorizontalScrollOnLoad(DependencyObject element)
        => (bool)element.GetValue(ResetHorizontalScrollOnLoadProperty);

    private static void OnResetHorizontalScrollOnLoadChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
    {
        if (d is not DataGrid dataGrid)
        {
            return;
        }

        dataGrid.Loaded -= DataGrid_Loaded;

        if ((bool)e.NewValue)
        {
            dataGrid.Loaded += DataGrid_Loaded;
        }
    }

    private static void DataGrid_Loaded(object sender, RoutedEventArgs e)
    {
        if (sender is not DataGrid dataGrid)
        {
            return;
        }

        dataGrid.Dispatcher.BeginInvoke(DispatcherPriority.Background, new Action(() =>
        {
            if (FindDescendant<ScrollViewer>(dataGrid) is { } scrollViewer)
            {
                scrollViewer.ScrollToHorizontalOffset(0);
            }
        }));
    }

    private static T? FindDescendant<T>(DependencyObject current) where T : DependencyObject
    {
        for (var i = 0; i < VisualTreeHelper.GetChildrenCount(current); i++)
        {
            var child = VisualTreeHelper.GetChild(current, i);
            if (child is T match)
            {
                return match;
            }

            var descendant = FindDescendant<T>(child);
            if (descendant is not null)
            {
                return descendant;
            }
        }

        return null;
    }
}
