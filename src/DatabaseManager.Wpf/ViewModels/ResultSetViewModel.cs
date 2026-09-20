using DatabaseManager.Core.Models;

namespace DatabaseManager.Wpf.ViewModels;

/// <summary>
/// Wraps one QueryResultSet for display. Currently only the header-text computation has been
/// pulled out of MainWindow's imperative Expander-building code (RenderResultsSections still
/// builds the actual DataGrid/Expander/ContextMenu there) - this is the seam Phase 9's
/// Procedure Runner extraction is expected to reuse once its results display moves too.
/// </summary>
public sealed class ResultSetViewModel(QueryResultSet resultSet)
{
    public QueryResultSet ResultSet { get; } = resultSet;

    public string Title => ResultSet.Title;

    public System.Data.DataTable? DataTable => ResultSet.DataTable;

    public int AffectedRows => ResultSet.AffectedRows;

    public string HeaderText => DataTable is null
        ? $"{Title} ({AffectedRows} affected rows)"
        : $"{Title} ({DataTable.Rows.Count} rows, {DataTable.Columns.Count} columns)";
}
