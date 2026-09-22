using System.Globalization;
using System.Windows;
using System.Windows.Data;

namespace DatabaseManager.Wpf.Converters;

/// <summary>
/// Bool -> Visibility, inverted (Visible when false). Used for a query document tab's title
/// TextBlock, which should be visible everywhere EXCEPT while IsRenaming is true - the negation
/// of DirtyIndicatorConverter's bool source, so it needs its own converter rather than a
/// ConverterParameter flag.
/// </summary>
public sealed class InverseBooleanToVisibilityConverter : IValueConverter
{
    public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        => value is true ? Visibility.Collapsed : Visibility.Visible;

    public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        => throw new NotSupportedException();
}
