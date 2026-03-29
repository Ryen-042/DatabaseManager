namespace DatabaseManager.Core.Models.Schema;

public sealed class ColumnSchemaInfo
{
    public required string SchemaName { get; init; }

    public required string TableName { get; init; }

    public required string ColumnName { get; init; }

    public required string DataType { get; init; }

    public int OrdinalPosition { get; init; }

    public bool IsNullable { get; init; }

    public bool IsIdentity { get; init; }

    public bool IsPrimaryKey { get; init; }

    public short MaxLength { get; init; }

    public byte Precision { get; init; }

    public byte Scale { get; init; }
}
