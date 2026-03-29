namespace DatabaseManager.Core.Models.Schema;

public sealed class StoredProcedureSchemaInfo
{
    public required string SchemaName { get; init; }

    public required string ProcedureName { get; init; }

    public string FullName => $"[{SchemaName}].[{ProcedureName}]";
}
