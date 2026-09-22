using System.Globalization;
using System.Windows.Data;
using System.Windows.Media;

namespace DatabaseManager.Wpf.Converters;

/// <summary>
/// Hex color string (e.g. "#E5484D") -> SolidColorBrush, for a query document tab's color
/// indicator strip. Null/empty/unparseable -> Transparent rather than collapsing the element, so
/// every tab keeps identical spacing regardless of whether it has a color set.
/// </summary>
public sealed class HexColorToBrushConverter : IValueConverter
{
    public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
    {
        if (value is string hex && !string.IsNullOrWhiteSpace(hex))
        {
            try
            {
                return new SolidColorBrush((Color)ColorConverter.ConvertFromString(hex));
            }
            catch (FormatException)
            {
                // Falls through to Transparent below.
            }
        }

        return Brushes.Transparent;
    }

    public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        => throw new NotSupportedException();
}
