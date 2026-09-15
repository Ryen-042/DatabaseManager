using System.Collections;
using System.Windows;
using System.Windows.Controls;
using DatabaseManager.Wpf.ViewModels;

namespace DatabaseManager.Wpf.Controls;

public partial class ToastHost : UserControl
{
    public static readonly DependencyProperty ToastsProperty = DependencyProperty.Register(
        nameof(Toasts),
        typeof(IEnumerable),
        typeof(ToastHost),
        new PropertyMetadata(null, OnToastsChanged));

    public ToastHost()
    {
        InitializeComponent();
    }

    public IEnumerable? Toasts
    {
        get => (IEnumerable?)GetValue(ToastsProperty);
        set => SetValue(ToastsProperty, value);
    }

    private static void OnToastsChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
    {
        ((ToastHost)d).ToastItemsControl.ItemsSource = (IEnumerable?)e.NewValue;
    }
}
