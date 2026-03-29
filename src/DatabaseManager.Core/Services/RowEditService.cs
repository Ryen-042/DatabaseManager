using System.Data;
using DatabaseManager.Core.Models.Editing;
using DatabaseManager.Core.Models.Schema;
using Microsoft.Data.SqlClient;

namespace DatabaseManager.Core.Services;

public sealed class RowEditService : IRowEditService
{
    public async Task<DataTable> LoadTopRowsAsync(
        string connectionString,
        string schemaName,
        string tableName,
        int topRows,
        string? filterExpression,
        int commandTimeoutSeconds,
        CancellationToken cancellationToken)
    {
        var sql = $"SELECT TOP (@topRows) * FROM [{EscapeIdentifier(schemaName)}].[{EscapeIdentifier(tableName)}]";
        if (!string.IsNullOrWhiteSpace(filterExpression))
        {
            sql += $" WHERE {filterExpression}";
        }

        await using var connection = new SqlConnection(connectionString);
        await connection.OpenAsync(cancellationToken);

        await using var command = new SqlCommand(sql, connection)
        {
            CommandTimeout = commandTimeoutSeconds
        };
        command.Parameters.Add(new SqlParameter("@topRows", SqlDbType.Int) { Value = topRows });

        await using var reader = await command.ExecuteReaderAsync(cancellationToken);
        var table = new DataTable();
        table.Load(reader);
        table.AcceptChanges();
        return table;
    }

    public async Task<int> SaveUpdatedRowsAsync(
        string connectionString,
        string schemaName,
        string tableName,
        IReadOnlyList<ColumnSchemaInfo> columns,
        IReadOnlyList<RowUpdateRequest> rowUpdates,
        int commandTimeoutSeconds,
        CancellationToken cancellationToken)
    {
        var primaryKeys = columns.Where(c => c.IsPrimaryKey).Select(c => c.ColumnName).ToList();
        if (primaryKeys.Count == 0)
        {
            throw new InvalidOperationException("Cannot save changes because the selected table has no primary key.");
        }

        var updatableColumns = columns
            .Where(c => !c.IsPrimaryKey && !c.IsIdentity)
            .Select(c => c.ColumnName)
            .ToList();

        if (updatableColumns.Count == 0 || rowUpdates.Count == 0)
        {
            return 0;
        }

        await using var connection = new SqlConnection(connectionString);
        await connection.OpenAsync(cancellationToken);
        await using var transaction = (SqlTransaction)await connection.BeginTransactionAsync(cancellationToken);

        try
        {
            var affectedRows = 0;

            foreach (var update in rowUpdates)
            {
                var setColumns = updatableColumns
                    .Where(name => update.CurrentValues.ContainsKey(name))
                    .ToList();

                if (setColumns.Count == 0)
                {
                    continue;
                }

                EnsurePrimaryKeyValues(primaryKeys, update.OriginalKeyValues);

                var setClause = string.Join(", ", setColumns.Select(c => $"[{EscapeIdentifier(c)}] = @set_{c}"));
                var whereClause = string.Join(" AND ", primaryKeys.Select(c => $"[{EscapeIdentifier(c)}] = @key_{c}"));
                var sql = $"UPDATE [{EscapeIdentifier(schemaName)}].[{EscapeIdentifier(tableName)}] SET {setClause} WHERE {whereClause};";

                await using var command = new SqlCommand(sql, connection, transaction)
                {
                    CommandTimeout = commandTimeoutSeconds
                };

                foreach (var column in setColumns)
                {
                    command.Parameters.AddWithValue($"@set_{column}", ToDbValue(update.CurrentValues[column]));
                }

                foreach (var key in primaryKeys)
                {
                    command.Parameters.AddWithValue($"@key_{key}", ToDbValue(update.OriginalKeyValues[key]));
                }

                affectedRows += await command.ExecuteNonQueryAsync(cancellationToken);
            }

            await transaction.CommitAsync(cancellationToken);
            return affectedRows;
        }
        catch
        {
            await transaction.RollbackAsync(cancellationToken);
            throw;
        }
    }

    public async Task<int> DeleteRowAsync(
        string connectionString,
        string schemaName,
        string tableName,
        IReadOnlyList<ColumnSchemaInfo> columns,
        IReadOnlyDictionary<string, object?> keyValues,
        int commandTimeoutSeconds,
        CancellationToken cancellationToken)
    {
        var primaryKeys = columns.Where(c => c.IsPrimaryKey).Select(c => c.ColumnName).ToList();
        if (primaryKeys.Count == 0)
        {
            throw new InvalidOperationException("Cannot delete rows because the selected table has no primary key.");
        }

        EnsurePrimaryKeyValues(primaryKeys, keyValues);

        var whereClause = string.Join(" AND ", primaryKeys.Select(c => $"[{EscapeIdentifier(c)}] = @key_{c}"));
        var sql = $"DELETE FROM [{EscapeIdentifier(schemaName)}].[{EscapeIdentifier(tableName)}] WHERE {whereClause};";

        await using var connection = new SqlConnection(connectionString);
        await connection.OpenAsync(cancellationToken);
        await using var transaction = (SqlTransaction)await connection.BeginTransactionAsync(cancellationToken);

        try
        {
            await using var command = new SqlCommand(sql, connection, transaction)
            {
                CommandTimeout = commandTimeoutSeconds
            };

            foreach (var key in primaryKeys)
            {
                command.Parameters.AddWithValue($"@key_{key}", ToDbValue(keyValues[key]));
            }

            var affectedRows = await command.ExecuteNonQueryAsync(cancellationToken);
            await transaction.CommitAsync(cancellationToken);
            return affectedRows;
        }
        catch
        {
            await transaction.RollbackAsync(cancellationToken);
            throw;
        }
    }

    private static string EscapeIdentifier(string value)
    {
        return value.Replace("]", "]]", StringComparison.Ordinal);
    }

    private static object ToDbValue(object? value)
    {
        return value ?? DBNull.Value;
    }

    private static void EnsurePrimaryKeyValues(IReadOnlyList<string> primaryKeys, IReadOnlyDictionary<string, object?> keyValues)
    {
        foreach (var key in primaryKeys)
        {
            if (!keyValues.TryGetValue(key, out var value) || value is null or DBNull)
            {
                throw new InvalidOperationException($"Primary key value is required for column '{key}'.");
            }
        }
    }
}
