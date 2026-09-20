using System.Globalization;
using System.Windows.Data;
using DatabaseManager.Core.Services;

namespace DatabaseManager.Wpf.Converters;

/// <summary>
/// Applied programmatically (not via XAML) to auto-generated Results/Procedure Runner grid
/// columns - see MainWindow's ResultsDataGrid_AutoGeneratingColumn. Always formats without
/// full-output mode; the grids that need full-output-aware formatting read
/// DisplayValueFormatter directly instead of going through a converter.
/// </summary>
public sealed class ResultValueConverter : IValueConverter
{
    public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        => DisplayValueFormatter.FormatForDisplay(value);

    public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        => value;
}
