using DatabaseManager.Core.Models.Schema;

namespace DatabaseManager.Core.Services.Schema;

public interface IDatabaseSchemaService
{
    Task<IReadOnlyList<TableSchemaInfo>> GetTablesAsync(string connectionString, CancellationToken cancellationToken);

    Task<IReadOnlyList<ForeignKeySchemaInfo>> GetForeignKeysAsync(string connectionString, CancellationToken cancellationToken);

    Task<IReadOnlyList<ColumnSchemaInfo>> GetColumnsAsync(string connectionString, string schemaName, string tableName, CancellationToken cancellationToken);

    Task<IReadOnlyList<StoredProcedureSchemaInfo>> GetStoredProceduresAsync(string connectionString, CancellationToken cancellationToken);

    Task<IReadOnlyList<StoredProcedureParameterInfo>> GetStoredProcedureParametersAsync(
        string connectionString,
        string schemaName,
        string procedureName,
        CancellationToken cancellationToken);

    Task<string?> GetStoredProcedureDefinitionAsync(
        string connectionString,
        string schemaName,
        string procedureName,
        CancellationToken cancellationToken);
}
