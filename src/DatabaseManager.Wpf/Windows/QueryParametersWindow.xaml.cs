using System.Collections.ObjectModel;
using System.Linq;
using System.Windows;
using DatabaseManager.Core.Models;

namespace DatabaseManager.Wpf.Windows;

public partial class QueryParametersWindow : Window
{
    private readonly ObservableCollection<QueryParameterEditorRow> _rows;

    public QueryParametersWindow(IReadOnlyList<string> parameterNames)
    {
        InitializeComponent();

        _rows = new ObservableCollection<QueryParameterEditorRow>(
            parameterNames.Select(name => new QueryParameterEditorRow
            {
                ParameterName = name,
                Value = string.Empty,
                SendAsNull = false
            }));

        ParametersDataGrid.ItemsSource = _rows;
    }

    public IReadOnlyList<QueryParameterValue> Parameters { get; private set; } = Array.Empty<QueryParameterValue>();

    private void ExecuteButton_Click(object sender, RoutedEventArgs e)
    {
        Parameters = _rows
            .Select(row => new QueryParameterValue
            {
                Name = row.ParameterName,
                Value = row.SendAsNull ? null : row.Value
            })
            .ToList();

        DialogResult = true;
    }
}

public sealed class QueryParameterEditorRow
{
    public required string ParameterName { get; init; }

    public string? Value { get; set; }

    public bool SendAsNull { get; set; }
}
