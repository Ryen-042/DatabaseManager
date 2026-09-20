using System.Data;
using DatabaseManager.Core.Models;
using DatabaseManager.Wpf.ViewModels;

namespace DatabaseManager.Tests.ViewModels;

public sealed class ResultSetViewModelTests
{
    [Fact]
    public void HeaderText_WithDataTable_ShowsRowAndColumnCounts()
    {
        var table = new DataTable();
        table.Columns.Add("Id");
        table.Columns.Add("Name");
        table.Rows.Add(1, "Alice");
        table.Rows.Add(2, "Bob");

        var vm = new ResultSetViewModel(new QueryResultSet { Title = "Result Set 1", DataTable = table, AffectedRows = 2 });

        Assert.Equal("Result Set 1 (2 rows, 2 columns)", vm.HeaderText);
    }

    [Fact]
    public void HeaderText_WithoutDataTable_ShowsAffectedRows()
    {
        var vm = new ResultSetViewModel(new QueryResultSet { Title = "Statement Summary", DataTable = null, AffectedRows = 5 });

        Assert.Equal("Statement Summary (5 affected rows)", vm.HeaderText);
    }

    [Fact]
    public void Properties_DelegateToWrappedResultSet()
    {
        var table = new DataTable();
        var resultSet = new QueryResultSet { Title = "Foo", DataTable = table, AffectedRows = 3 };

        var vm = new ResultSetViewModel(resultSet);

        Assert.Same(resultSet, vm.ResultSet);
        Assert.Equal("Foo", vm.Title);
        Assert.Same(table, vm.DataTable);
        Assert.Equal(3, vm.AffectedRows);
    }
}
