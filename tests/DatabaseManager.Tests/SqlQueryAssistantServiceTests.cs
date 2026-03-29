using DatabaseManager.Core.Models.Schema;
using DatabaseManager.Core.Services.Schema;

namespace DatabaseManager.Tests;

public sealed class SqlQueryAssistantServiceTests
{
    private readonly SqlQueryAssistantService _service = new();

    [Fact]
    public void BuildSelectTopQuery_ReturnsExpectedSql()
    {
        var table = new TableSchemaInfo
        {
            SchemaName = "dbo",
            TableName = "Customers",
            RowCount = 0
        };

        var sql = _service.BuildSelectTopQuery(table);

        Assert.Contains("SELECT TOP (100)", sql);
        Assert.Contains("[dbo].[Customers]", sql);
    }

    [Fact]
    public void BuildUpdateQuery_WithoutPrimaryKey_AddsSafeTodoWhereClause()
    {
        var table = new TableSchemaInfo
        {
            SchemaName = "sales",
            TableName = "Orders",
            RowCount = 0
        };

        var columns = new List<ColumnSchemaInfo>
        {
            new()
            {
                SchemaName = "sales",
                TableName = "Orders",
                ColumnName = "OrderId",
                DataType = "int",
                OrdinalPosition = 1,
                IsNullable = false,
                IsIdentity = true,
                IsPrimaryKey = false,
                MaxLength = 4,
                Precision = 10,
                Scale = 0
            },
            new()
            {
                SchemaName = "sales",
                TableName = "Orders",
                ColumnName = "Status",
                DataType = "nvarchar",
                OrdinalPosition = 2,
                IsNullable = false,
                IsIdentity = false,
                IsPrimaryKey = false,
                MaxLength = 40,
                Precision = 0,
                Scale = 0
            }
        };

        var sql = _service.BuildUpdateQuery(table, columns);

        Assert.Contains("TODO", sql);
        Assert.Contains("1 = 0", sql);
    }
}
