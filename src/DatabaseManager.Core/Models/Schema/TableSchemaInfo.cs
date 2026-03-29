namespace DatabaseManager.Core.Models.Schema;

public sealed class TableSchemaInfo
{
    public required string SchemaName { get; init; }

    public required string TableName { get; init; }

    public long RowCount { get; init; }

    public string FullName => $"[{SchemaName}].[{TableName}]";
}
