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

    /// <summary>
    /// Saves updates (by primary key, with an optimistic-concurrency check against the row's
    /// original values - see RowEditSqlBuilder.BuildConcurrencyMatchColumns - when the table has
    /// one) and inserts transactionally. WasExecuted is false and nothing is committed if any
    /// update hits a concurrency conflict; see RowSaveResult.
    /// </summary>
    Task<RowSaveResult> SaveRowChangesAsync(
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

    Task<RowDeleteByColumnsResult> DeleteRowsBySelectedColumnsAsync(
        string connectionString,
        string schemaName,
        string tableName,
        IReadOnlyList<string> selectedColumns,
        IReadOnlyList<IReadOnlyDictionary<string, object?>> selectedRows,
        int commandTimeoutSeconds,
        CancellationToken cancellationToken);
}
