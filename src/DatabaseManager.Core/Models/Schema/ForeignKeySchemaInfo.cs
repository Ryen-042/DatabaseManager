namespace DatabaseManager.Core.Models.Schema;

public sealed class ForeignKeySchemaInfo
{
    public required string ConstraintName { get; init; }

    public required string ParentSchemaName { get; init; }

    public required string ParentTableName { get; init; }

    public required string ParentColumnName { get; init; }

    public required string ReferencedSchemaName { get; init; }

    public required string ReferencedTableName { get; init; }

    public required string ReferencedColumnName { get; init; }
}
