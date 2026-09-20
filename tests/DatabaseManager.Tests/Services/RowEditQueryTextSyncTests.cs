using DatabaseManager.Core.Services;

namespace DatabaseManager.Tests.Services;

public sealed class RowEditQueryTextSyncTests
{
    [Fact]
    public void BuildQuery_NoFilterOrOrderBy_ProducesSimpleTopQuery()
    {
        var sql = RowEditQueryTextSync.BuildQuery("dbo", "Users", 200, filter: null, orderBy: null);

        Assert.Equal($"SELECT TOP (200) *{Environment.NewLine}FROM [dbo].[Users];", sql);
    }

    [Fact]
    public void BuildQuery_WithFilterAndOrderBy_AppendsBothClauses()
    {
        var sql = RowEditQueryTextSync.BuildQuery("dbo", "Users", 50, filter: "Age > 18", orderBy: "Name DESC");

        Assert.Equal(
            $"SELECT TOP (50) *{Environment.NewLine}FROM [dbo].[Users]{Environment.NewLine}WHERE Age > 18{Environment.NewLine}ORDER BY Name DESC;",
            sql);
    }

    [Fact]
    public void BuildQuery_WhitespaceOnlyFilterAndOrderBy_TreatedAsAbsent()
    {
        var sql = RowEditQueryTextSync.BuildQuery("dbo", "Users", 100, filter: "   ", orderBy: "  ");

        Assert.Equal($"SELECT TOP (100) *{Environment.NewLine}FROM [dbo].[Users];", sql);
    }

    [Fact]
    public void TryParse_RoundTripsWhatBuildQueryProduces()
    {
        var sql = RowEditQueryTextSync.BuildQuery("dbo", "Users", 75, "Age > 18", "Name DESC");

        var success = RowEditQueryTextSync.TryParse(sql, out var topRows, out var filter, out var orderBy);

        Assert.True(success);
        Assert.Equal(75, topRows);
        Assert.Equal("Age > 18", filter);
        Assert.Equal("Name DESC", orderBy);
    }

    [Fact]
    public void TryParse_SimpleQueryWithoutTopFilterOrOrderBy_DefaultsTopTo200()
    {
        var success = RowEditQueryTextSync.TryParse("SELECT * FROM [dbo].[Users]", out var topRows, out var filter, out var orderBy);

        Assert.True(success);
        Assert.Equal(200, topRows);
        Assert.Null(filter);
        Assert.Null(orderBy);
    }

    [Fact]
    public void TryParse_IsCaseInsensitiveAndToleratesExtraWhitespace()
    {
        var success = RowEditQueryTextSync.TryParse(
            "  select   top(10)   *   from   [dbo].[Users]   where Age > 18   order by Name  ",
            out var topRows, out var filter, out var orderBy);

        Assert.True(success);
        Assert.Equal(10, topRows);
        Assert.Equal("Age > 18", filter);
        Assert.Equal("Name", orderBy);
    }

    [Fact]
    public void TryParse_TrailingSemicolonOnOrderBy_IsStripped()
    {
        var success = RowEditQueryTextSync.TryParse(
            "SELECT TOP (10) * FROM [dbo].[Users] ORDER BY Name;",
            out _, out _, out var orderBy);

        Assert.True(success);
        Assert.Equal("Name", orderBy);
    }

    [Theory]
    [InlineData("")]
    [InlineData("-- a comment\nSELECT * FROM [dbo].[Users]")]
    [InlineData("SELECT Name FROM [dbo].[Users]")] // not a "SELECT *" query
    [InlineData("UPDATE [dbo].[Users] SET Name = 'x'")]
    public void TryParse_NonMatchingText_ReturnsFalse(string sql)
    {
        var success = RowEditQueryTextSync.TryParse(sql, out _, out _, out _);

        Assert.False(success);
    }

    [Fact]
    public void TryParse_ZeroOrNegativeTop_ReturnsFalse()
    {
        var success = RowEditQueryTextSync.TryParse("SELECT TOP (0) * FROM [dbo].[Users]", out _, out _, out _);

        Assert.False(success);
    }
}
