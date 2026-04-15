using DatabaseManager.Core.Services;

namespace DatabaseManager.Tests;

public sealed class QueryOutputModeParserTests
{
    [Fact]
    public void Parse_WithTrailingFullDirective_EnablesFullOutputAndSanitizesSql()
    {
        var sql = "SELECT 1;\n-- full";

        var result = QueryOutputModeParser.Parse(sql);

        Assert.True(result.FullOutputEnabled);
        Assert.True(result.HasFullDirective);
        Assert.Equal("SELECT 1;", result.Sql);
    }

    [Fact]
    public void Parse_WithTrailingCompactFullDirective_EnablesFullOutputAndSanitizesSql()
    {
        var sql = "SELECT 1;\n--full";

        var result = QueryOutputModeParser.Parse(sql);

        Assert.True(result.FullOutputEnabled);
        Assert.True(result.HasFullDirective);
        Assert.Equal("SELECT 1;", result.Sql);
    }

    [Fact]
    public void Parse_WithNonTrailingComment_DoesNotEnableFullOutput()
    {
        var sql = "SELECT 1;\n-- full\nSELECT 2;";

        var result = QueryOutputModeParser.Parse(sql);

        Assert.False(result.FullOutputEnabled);
        Assert.False(result.HasFullDirective);
        Assert.Equal(sql, result.Sql);
    }

    [Fact]
    public void Parse_WithDifferentTrailingComment_DoesNotEnableFullOutput()
    {
        var sql = "SELECT 1;\n-- debug";

        var result = QueryOutputModeParser.Parse(sql);

        Assert.False(result.FullOutputEnabled);
        Assert.False(result.HasFullDirective);
        Assert.Equal(sql, result.Sql);
    }

    [Fact]
    public void Parse_IgnoresCommentMarkersInsideStringLiteral()
    {
        var sql = "SELECT '-- full' AS Value;";

        var result = QueryOutputModeParser.Parse(sql);

        Assert.False(result.FullOutputEnabled);
        Assert.False(result.HasFullDirective);
        Assert.Equal(sql, result.Sql);
    }

    [Fact]
    public void ExtractParameterNames_ReturnsDistinctNamesInOrder()
    {
        var sql = "SELECT * FROM Users WHERE UserId = @userId AND StartDate >= @startDate AND UserId <> @userId;";

        var names = QueryOutputModeParser.ExtractParameterNames(sql);

        Assert.Equal(new[] { "@userId", "@startDate" }, names);
    }

    [Fact]
    public void ExtractParameterNames_IgnoresStringLiteralsAndComments()
    {
        var sql = "SELECT '@fake', @real -- @ignore\n/* @ignoreToo */ WHERE Id = @real2";

        var names = QueryOutputModeParser.ExtractParameterNames(sql);

        Assert.Equal(new[] { "@real", "@real2" }, names);
    }
}
