using System.Globalization;
using System.Windows;
using System.Windows.Data;

namespace DatabaseManager.Wpf.Converters;

/// <summary>
/// Bool -> Visibility for the dirty dot on a query document's tab header (Visible when dirty).
/// </summary>
public sealed class DirtyIndicatorConverter : IValueConverter
{
    public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        => value is true ? Visibility.Visible : Visibility.Collapsed;

    public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        => throw new NotSupportedException();
}
