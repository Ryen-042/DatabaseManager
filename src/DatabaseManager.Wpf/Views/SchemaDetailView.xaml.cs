using System.Windows.Controls;
using System.Windows.Input;

namespace DatabaseManager.Wpf.Views;

public partial class SchemaDetailView : UserControl
{
    public SchemaDetailView() => InitializeComponent();

    private void SchemaDataGrid_PreviewMouseLeftButtonUp(object sender, MouseButtonEventArgs e)
    {
        // Double-click copy behavior is deliberately disabled here, matching every other grid.
        e.Handled = false;
    }
}
