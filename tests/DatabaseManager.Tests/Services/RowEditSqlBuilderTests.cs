using DatabaseManager.Core.Services;

namespace DatabaseManager.Tests.Services;

public sealed class RowEditSqlBuilderTests
{
    [Fact]
    public void EscapeIdentifier_DoublesClosingBrackets()
    {
        Assert.Equal("Weird]]Name", RowEditSqlBuilder.EscapeIdentifier("Weird]Name"));
    }

    [Fact]
    public void BuildUpdateStatement_WithPrimaryKey_UsesNullSafeWhereAndNoRowCountGuard()
    {
        var sql = RowEditSqlBuilder.BuildUpdateStatement(
            "dbo", "Users",
            setColumns: ["Name", "Email"],
            matchColumns: ["Id"],
            hasPrimaryKey: true);

        Assert.Equal(
            "UPDATE [dbo].[Users] SET [Name] = @set_Name, [Email] = @set_Email " +
            "WHERE ((@key_Id IS NULL AND [Id] IS NULL) OR [Id] = @key_Id);",
            sql);
    }

    [Fact]
    public void BuildUpdateStatement_WithoutPrimaryKey_AppendsRowCountGuard()
    {
        var sql = RowEditSqlBuilder.BuildUpdateStatement(
            "dbo", "Users",
            setColumns: ["Name"],
            matchColumns: ["Name", "Email"],
            hasPrimaryKey: false);

        Assert.Contains("IF @@ROWCOUNT <> 1", sql);
        Assert.Contains("THROW 50000, 'Cannot save changes because the row is not uniquely identifiable without a primary key.', 1;", sql);
    }

    [Fact]
    public void BuildInsertStatement_ListsColumnsAndParameters()
    {
        var sql = RowEditSqlBuilder.BuildInsertStatement("dbo", "Users", ["Name", "Email"]);

        Assert.Equal("INSERT INTO [dbo].[Users] ([Name], [Email]) VALUES (@ins_Name, @ins_Email);", sql);
    }

    [Fact]
    public void BuildDeleteByPrimaryKeyStatement_UsesExactMatchOnEachKeyColumn()
    {
        var sql = RowEditSqlBuilder.BuildDeleteByPrimaryKeyStatement("dbo", "OrderItems", ["OrderId", "LineNo"]);

        Assert.Equal("DELETE FROM [dbo].[OrderItems] WHERE [OrderId] = @key_OrderId AND [LineNo] = @key_LineNo;", sql);
    }

    [Fact]
    public void BuildCountSql_UsesNullSafePredicateForEachSelectedColumn()
    {
        var sql = RowEditSqlBuilder.BuildCountSql("dbo", "Users", ["Name", "Email"]);

        Assert.Equal(
            "SELECT COUNT(1) FROM [dbo].[Users] WHERE " +
            "((@Name IS NULL AND [Name] IS NULL) OR [Name] = @Name) AND " +
            "((@Email IS NULL AND [Email] IS NULL) OR [Email] = @Email);",
            sql);
    }

    [Fact]
    public void BuildDeleteTopOneSql_DeletesAtMostOneMatchingRow()
    {
        var sql = RowEditSqlBuilder.BuildDeleteTopOneSql("dbo", "Users", ["Name"]);

        Assert.StartsWith("DELETE TOP (1) FROM [dbo].[Users] WHERE", sql);
    }

    [Fact]
    public void BuildGeneratedDeleteSql_EmitsOneDeleteStatementPerRowWithLiteralValues()
    {
        var rows = new List<IReadOnlyDictionary<string, object?>>
        {
            new Dictionary<string, object?> { ["Name"] = "Alice", ["Age"] = null }
        };

        var sql = RowEditSqlBuilder.BuildGeneratedDeleteSql("dbo", "Users", ["Name", "Age"], rows);

        Assert.Contains("DELETE TOP (1) FROM [dbo].[Users]", sql);
        Assert.Contains("[Name] = N'Alice'", sql);
        Assert.Contains("[Age] IS NULL", sql);
        Assert.Contains("GO", sql);
    }

    [Fact]
    public void BuildGeneratedDeleteSql_MissingColumnValue_Throws()
    {
        var rows = new List<IReadOnlyDictionary<string, object?>>
        {
            new Dictionary<string, object?>()
        };

        Assert.Throws<InvalidOperationException>(() =>
            RowEditSqlBuilder.BuildGeneratedDeleteSql("dbo", "Users", ["Name"], rows));
    }

    [Theory]
    [InlineData("O'Brien", "N'O''Brien'")]
    public void ToSqlLiteral_String_EscapesQuotes(string input, string expected)
    {
        Assert.Equal(expected, RowEditSqlBuilder.ToSqlLiteral(input));
    }

    [Fact]
    public void ToSqlLiteral_Bool_RendersAsOneOrZero()
    {
        Assert.Equal("1", RowEditSqlBuilder.ToSqlLiteral(true));
        Assert.Equal("0", RowEditSqlBuilder.ToSqlLiteral(false));
    }

    [Fact]
    public void ToSqlLiteral_ByteArray_RendersAsHexLiteral()
    {
        Assert.Equal("0x0A1B", RowEditSqlBuilder.ToSqlLiteral(new byte[] { 0x0A, 0x1B }));
    }

    [Fact]
    public void ToSqlLiteral_Guid_RendersQuoted()
    {
        var guid = Guid.Parse("11111111-2222-3333-4444-555555555555");

        Assert.Equal("'11111111-2222-3333-4444-555555555555'", RowEditSqlBuilder.ToSqlLiteral(guid));
    }

    [Fact]
    public void ToSqlLiteral_Int_RendersInvariantly()
    {
        Assert.Equal("42", RowEditSqlBuilder.ToSqlLiteral(42));
    }
}
