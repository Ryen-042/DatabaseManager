using System.Collections.ObjectModel;
using System.Linq;
using System.Windows;
using DatabaseManager.Core.Models.Schema;

namespace DatabaseManager.Wpf.Windows;

public partial class DeleteColumnsSelectionWindow : Window
{
    private readonly ObservableCollection<DeleteColumnSelectionRow> _rows;

    public DeleteColumnsSelectionWindow(IReadOnlyList<ColumnSchemaInfo> availableColumns)
    {
        InitializeComponent();

        _rows = new ObservableCollection<DeleteColumnSelectionRow>(
            availableColumns
                .OrderBy(c => c.OrdinalPosition)
                .Select(c => new DeleteColumnSelectionRow
                {
                    ColumnName = c.ColumnName,
                    DataType = c.DataType,
                    IsSelected = c.IsPrimaryKey
                }));

        ColumnsDataGrid.ItemsSource = _rows;
    }

    public IReadOnlyList<string> SelectedColumns { get; private set; } = Array.Empty<string>();

    private void ApplyButton_Click(object sender, RoutedEventArgs e)
    {
        SelectedColumns = _rows
            .Where(row => row.IsSelected)
            .Select(row => row.ColumnName)
            .ToList();

        DialogResult = true;
    }
}

public sealed class DeleteColumnSelectionRow
{
    public required string ColumnName { get; init; }

    public required string DataType { get; init; }

    public bool IsSelected { get; set; }
}
