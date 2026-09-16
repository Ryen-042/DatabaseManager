using System.Windows.Controls;
using System.Windows.Input;
using DatabaseManager.Wpf.ViewModels;

namespace DatabaseManager.Wpf.Views;

public partial class TemplatesPanelView : UserControl
{
    public TemplatesPanelView()
    {
        InitializeComponent();
    }

    private void TemplatesListBox_MouseDoubleClick(object sender, MouseButtonEventArgs e)
    {
        (DataContext as TemplatesPanelViewModel)?.ActivateSelected();
    }
}
