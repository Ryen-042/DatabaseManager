using DatabaseManager.Core.Models.Schema;
using DatabaseManager.Core.Services.Schema;
using DatabaseManager.Wpf.SqlSuggestions;

namespace DatabaseManager.Tests;

public sealed class SqlQueryAssistantServiceTests
{
    private readonly SqlQueryAssistantService _service = new();
    private readonly SqlSuggestionEngine _suggestionEngine = new();

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

    [Fact]
    public void BuildDropTableScript_UsesObjectIdGuard()
    {
        var table = new TableSchemaInfo
        {
            SchemaName = "dbo",
            TableName = "Customers",
            RowCount = 0
        };

        var sql = _service.BuildDropTableScript(table);

        Assert.Contains("OBJECT_ID", sql);
        Assert.Contains("DROP TABLE [dbo].[Customers]", sql);
    }

    [Fact]
    public void BuildAlterProcedureScript_ReplacesCreateWithAlter()
    {
        var procedure = new StoredProcedureSchemaInfo
        {
            SchemaName = "dbo",
            ProcedureName = "GetCustomers"
        };

        var definition = "CREATE PROCEDURE [dbo].[GetCustomers]\nAS\nBEGIN\nSELECT 1;\nEND";
        var sql = _service.BuildAlterProcedureScript(procedure, definition);

        Assert.Contains("ALTER PROCEDURE", sql);
        Assert.DoesNotContain("CREATE PROCEDURE", sql, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void GetSuggestions_WhenAliasIsQualified_PrioritizesColumnsFromMatchedTable()
    {
        var sql = "SELECT * FROM dbo.Customers c WHERE c.C";
        var token = "c.C";

        var suggestions = _suggestionEngine.GetSuggestions(new SqlSuggestionRequest
        {
            SqlText = sql,
            Token = token,
            TokenStart = sql.IndexOf(token, StringComparison.Ordinal),
            Tables = new[]
            {
                new TableSchemaInfo { SchemaName = "dbo", TableName = "Customers", RowCount = 0 },
                new TableSchemaInfo { SchemaName = "sales", TableName = "Orders", RowCount = 0 }
            },
            SelectedColumns = new[]
            {
                new ColumnSchemaInfo
                {
                    SchemaName = "dbo",
                    TableName = "Customers",
                    ColumnName = "CustomerId",
                    DataType = "int"
                },
                new ColumnSchemaInfo
                {
                    SchemaName = "dbo",
                    TableName = "Customers",
                    ColumnName = "CustomerName",
                    DataType = "nvarchar"
                },
                new ColumnSchemaInfo
                {
                    SchemaName = "sales",
                    TableName = "Orders",
                    ColumnName = "OrderId",
                    DataType = "int"
                }
            },
            StoredProcedures = Array.Empty<StoredProcedureSchemaInfo>(),
            ProcedureParameters = Array.Empty<StoredProcedureParameterInfo>(),
            ForeignKeys = Array.Empty<ForeignKeySchemaInfo>(),
            MaxResults = 5
        });

        Assert.Equal(new[] { "CustomerId", "CustomerName" }, suggestions.Take(2));
    }

    [Fact]
    public void GetSuggestions_WhenSchemaIsQualified_ReturnsMatchingTableFirst()
    {
        var sql = "SELECT * FROM dbo.C";
        var token = "dbo.C";

        var suggestions = _suggestionEngine.GetSuggestions(new SqlSuggestionRequest
        {
            SqlText = sql,
            Token = token,
            TokenStart = sql.IndexOf(token, StringComparison.Ordinal),
            Tables = new[]
            {
                new TableSchemaInfo { SchemaName = "dbo", TableName = "Customers", RowCount = 0 },
                new TableSchemaInfo { SchemaName = "sales", TableName = "Orders", RowCount = 0 }
            },
            SelectedColumns = Array.Empty<ColumnSchemaInfo>(),
            StoredProcedures = Array.Empty<StoredProcedureSchemaInfo>(),
            ProcedureParameters = Array.Empty<StoredProcedureParameterInfo>(),
            ForeignKeys = Array.Empty<ForeignKeySchemaInfo>(),
            MaxResults = 5
        });

        Assert.Equal("Customers", suggestions[0]);
        Assert.DoesNotContain("Orders", suggestions);
    }

    [Fact]
    public void GetSuggestions_WhenJoinContextHasForeignKey_ReturnsJoinHint()
    {
        var sql = "SELECT * FROM dbo.Customers c JOIN ";

        var suggestions = _suggestionEngine.GetSuggestions(new SqlSuggestionRequest
        {
            SqlText = sql,
            Token = string.Empty,
            TokenStart = sql.Length,
            Tables = new[]
            {
                new TableSchemaInfo { SchemaName = "dbo", TableName = "Customers", RowCount = 0 },
                new TableSchemaInfo { SchemaName = "sales", TableName = "Orders", RowCount = 0 }
            },
            SelectedColumns = new[]
            {
                new ColumnSchemaInfo
                {
                    SchemaName = "dbo",
                    TableName = "Customers",
                    ColumnName = "Id",
                    DataType = "int"
                },
                new ColumnSchemaInfo
                {
                    SchemaName = "sales",
                    TableName = "Orders",
                    ColumnName = "CustomerId",
                    DataType = "int"
                }
            },
            StoredProcedures = Array.Empty<StoredProcedureSchemaInfo>(),
            ProcedureParameters = Array.Empty<StoredProcedureParameterInfo>(),
            ForeignKeys = new[]
            {
                new ForeignKeySchemaInfo
                {
                    ConstraintName = "FK_Orders_Customers",
                    ParentSchemaName = "sales",
                    ParentTableName = "Orders",
                    ParentColumnName = "CustomerId",
                    ReferencedSchemaName = "dbo",
                    ReferencedTableName = "Customers",
                    ReferencedColumnName = "Id"
                }
            },
            MaxResults = 20
        });

        Assert.Contains(suggestions, suggestion =>
            suggestion.StartsWith("JOIN ", StringComparison.OrdinalIgnoreCase)
            && suggestion.Contains("[sales].[Orders]", StringComparison.OrdinalIgnoreCase)
            && suggestion.Contains("CustomerId", StringComparison.OrdinalIgnoreCase)
            && suggestion.Contains("[Id]", StringComparison.OrdinalIgnoreCase));
    }
}
