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

    [Theory]
    [InlineData("CREATE PROCEDURE dbo.MyProc @Id INT, @Name NVARCHAR(50) AS BEGIN SELECT 1; END;")]
    [InlineData("CREATE OR ALTER PROCEDURE dbo.MyProc @Id INT AS BEGIN SELECT 1; END;")]
    [InlineData("ALTER PROCEDURE dbo.MyProc @Id INT AS BEGIN SELECT 1; END;")]
    [InlineData("  create   proc dbo.MyProc @Id int as begin select 1; end;")]
    [InlineData("CREATE FUNCTION dbo.MyFunc(@Id INT) RETURNS INT AS BEGIN RETURN @Id; END;")]
    [InlineData("CREATE TRIGGER dbo.MyTrigger ON dbo.MyTable AFTER INSERT AS BEGIN SELECT @@ROWCOUNT; END;")]
    public void ExtractParameterNames_ReturnsEmptyForCreateOrAlterRoutineStatements(string sql)
    {
        var names = QueryOutputModeParser.ExtractParameterNames(sql);

        Assert.Empty(names);
    }

    [Fact]
    public void ExtractParameterNames_StillExtractsParametersForRegularQueries()
    {
        var sql = "UPDATE dbo.Users SET Name = @name WHERE Id = @id;";

        var names = QueryOutputModeParser.ExtractParameterNames(sql);

        Assert.Equal(new[] { "@name", "@id" }, names);
    }

    [Theory]
    [InlineData("CREATE PROCEDURE dbo.MyProc @Id INT AS BEGIN SELECT 1; END;", true)]
    [InlineData("CREATE OR ALTER PROCEDURE dbo.MyProc AS BEGIN SELECT 1; END;", true)]
    [InlineData("ALTER PROC dbo.MyProc AS BEGIN SELECT 1; END;", true)]
    [InlineData("SELECT * FROM dbo.Users WHERE Id = @id;", false)]
    [InlineData("CREATE TABLE dbo.Users (Id INT);", false)]
    public void IsCreateOrAlterRoutineStatement_DetectsRoutineDefinitions(string sql, bool expected)
    {
        Assert.Equal(expected, QueryOutputModeParser.IsCreateOrAlterRoutineStatement(sql));
    }
}
