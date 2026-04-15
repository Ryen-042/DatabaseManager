using System.Data;
using DatabaseManager.Core.Models.Editing;
using DatabaseManager.Core.Models.Schema;

namespace DatabaseManager.Core.Services;

public interface IRowEditService
{
    Task<DataTable> LoadTopRowsAsync(
        string connectionString,
        string schemaName,
        string tableName,
        int topRows,
        string? filterExpression,
        string? orderByExpression,
        int commandTimeoutSeconds,
        CancellationToken cancellationToken);

    Task<int> SaveUpdatedRowsAsync(
        string connectionString,
        string schemaName,
        string tableName,
        IReadOnlyList<ColumnSchemaInfo> columns,
        IReadOnlyList<RowUpdateRequest> rowUpdates,
        int commandTimeoutSeconds,
        CancellationToken cancellationToken);

    Task<int> SaveRowChangesAsync(
        string connectionString,
        string schemaName,
        string tableName,
        IReadOnlyList<ColumnSchemaInfo> columns,
        IReadOnlyList<RowUpdateRequest> rowUpdates,
        IReadOnlyList<RowInsertRequest> rowInserts,
        int commandTimeoutSeconds,
        CancellationToken cancellationToken);

    Task<int> DeleteRowAsync(
        string connectionString,
        string schemaName,
        string tableName,
        IReadOnlyList<ColumnSchemaInfo> columns,
        IReadOnlyDictionary<string, object?> keyValues,
        int commandTimeoutSeconds,
        CancellationToken cancellationToken);
}
