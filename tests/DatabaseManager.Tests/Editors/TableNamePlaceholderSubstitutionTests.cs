using DatabaseManager.Wpf.Editors;

namespace DatabaseManager.Tests.Editors;

public sealed class TableNamePlaceholderSubstitutionTests
{
    [Fact]
    public void TrySubstitute_BareTableName_IsReplaced()
    {
        var result = TableNamePlaceholderSubstitution.TrySubstitute("SELECT * FROM TableName", "dbo", "Users");

        Assert.Equal("SELECT * FROM [dbo].[Users]", result);
    }

    [Fact]
    public void TrySubstitute_BracketedTableName_IsReplaced()
    {
        var result = TableNamePlaceholderSubstitution.TrySubstitute("SELECT * FROM [TableName]", "dbo", "Users");

        Assert.Equal("SELECT * FROM [dbo].[Users]", result);
    }

    [Fact]
    public void TrySubstitute_SchemaQualifiedBracketedPlaceholder_ReplacesWholeReference()
    {
        // This is exactly the shape of the built-in SQL suggestion snippets
        // (e.g. "SELECT TOP (100) * FROM [dbo].[TableName];") - the hardcoded
        // "dbo" schema must be replaced too, not left behind as a stray prefix.
        var result = TableNamePlaceholderSubstitution.TrySubstitute("SELECT TOP (100) * FROM [dbo].[TableName];", "sales", "Orders");

        Assert.Equal("SELECT TOP (100) * FROM [sales].[Orders];", result);
    }

    [Fact]
    public void TrySubstitute_SchemaQualifiedUnbracketedPlaceholder_ReplacesWholeReference()
    {
        var result = TableNamePlaceholderSubstitution.TrySubstitute("SELECT * FROM dbo.TableName", "sales", "Orders");

        Assert.Equal("SELECT * FROM [sales].[Orders]", result);
    }

    [Fact]
    public void TrySubstitute_MultipleOccurrences_ReplacesAll()
    {
        var result = TableNamePlaceholderSubstitution.TrySubstitute(
            "INSERT INTO [dbo].[TableName] SELECT * FROM TableName",
            "dbo",
            "Users");

        Assert.Equal("INSERT INTO [dbo].[Users] SELECT * FROM [dbo].[Users]", result);
    }

    [Fact]
    public void TrySubstitute_NoPlaceholderPresent_ReturnsNull()
    {
        var result = TableNamePlaceholderSubstitution.TrySubstitute("SELECT * FROM [dbo].[Users]", "dbo", "Users");

        Assert.Null(result);
    }

    [Fact]
    public void TrySubstitute_CaseSensitive_DoesNotMatchDifferentCasing()
    {
        var result = TableNamePlaceholderSubstitution.TrySubstitute("SELECT * FROM tablename", "dbo", "Users");

        Assert.Null(result);
    }
}
