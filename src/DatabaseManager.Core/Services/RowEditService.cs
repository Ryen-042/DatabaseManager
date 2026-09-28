using System.Data;
using System.Globalization;
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
        string? orderByExpression,
        int commandTimeoutSeconds,
        CancellationToken cancellationToken)
    {
        var sql = $"SELECT TOP (@topRows) * FROM [{RowEditSqlBuilder.EscapeIdentifier(schemaName)}].[{RowEditSqlBuilder.EscapeIdentifier(tableName)}]";
        if (!string.IsNullOrWhiteSpace(filterExpression))
        {
            sql += $" WHERE {filterExpression}";
        }

        if (!string.IsNullOrWhiteSpace(orderByExpression))
        {
            sql += $" ORDER BY {orderByExpression}";
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

    public async Task<RowSaveResult> SaveRowChangesAsync(
        string connectionString,
        string schemaName,
        string tableName,
        IReadOnlyList<ColumnSchemaInfo> columns,
        IReadOnlyList<RowUpdateRequest> rowUpdates,
        IReadOnlyList<RowInsertRequest> rowInserts,
        int commandTimeoutSeconds,
        CancellationToken cancellationToken)
    {
        var primaryKeys = columns.Where(c => c.IsPrimaryKey).Select(c => c.ColumnName).ToList();
        var hasPrimaryKey = primaryKeys.Count > 0;

        // With a PK, matchColumns also carries the optimistic-concurrency check (a rowversion
        // column alone, or every other comparable column) alongside the identity predicate; the
        // no-PK path is unchanged - it already matches on every column and relies on its own
        // THROW-guarded statement (BuildUpdateStatement's hasPrimaryKey: false branch) instead.
        var matchColumns = hasPrimaryKey
            ? RowEditSqlBuilder.BuildConcurrencyMatchColumns(columns)
            : columns.Select(c => c.ColumnName).ToList();

        var updatableColumns = columns
            .Where(c => !c.IsPrimaryKey && !c.IsIdentity && !RowEditSqlBuilder.IsRowVersionColumn(c))
            .Select(c => c.ColumnName)
            .ToList();

        var insertableColumns = columns
            .Where(c => !c.IsIdentity && !RowEditSqlBuilder.IsRowVersionColumn(c))
            .Select(c => c.ColumnName)
            .ToList();

        var totalIntendedChanges = rowUpdates.Count + rowInserts.Count;

        if ((updatableColumns.Count == 0 || rowUpdates.Count == 0) && rowInserts.Count == 0)
        {
            return new RowSaveResult
            {
                TotalIntendedChanges = totalIntendedChanges,
                AffectedRows = 0,
                WasExecuted = true,
                ConflictedRowKeys = []
            };
        }

        await using var connection = new SqlConnection(connectionString);
        await connection.OpenAsync(cancellationToken);
        await using var transaction = (SqlTransaction)await connection.BeginTransactionAsync(cancellationToken);

        try
        {
            var affectedRows = 0;
            var conflictedRowKeys = new List<IReadOnlyDictionary<string, object?>>();

            foreach (var update in rowUpdates)
            {
                var setColumns = updatableColumns
                    .Where(name => update.CurrentValues.ContainsKey(name))
                    .ToList();

                if (setColumns.Count == 0)
                {
                    continue;
                }

                EnsureMatchValues(matchColumns, update.OriginalKeyValues);

                var sql = RowEditSqlBuilder.BuildUpdateStatement(schemaName, tableName, setColumns, matchColumns, hasPrimaryKey);

                await using var command = new SqlCommand(sql, connection, transaction)
                {
                    CommandTimeout = commandTimeoutSeconds
                };

                foreach (var column in setColumns)
                {
                    command.Parameters.AddWithValue($"@set_{column}", ToDbValue(update.CurrentValues[column]));
                }

                foreach (var key in matchColumns)
                {
                    command.Parameters.AddWithValue($"@key_{key}", ToDbValue(update.OriginalKeyValues[key]));
                }

                var rowsAffected = await command.ExecuteNonQueryAsync(cancellationToken);

                if (hasPrimaryKey && rowsAffected == 0)
                {
                    // Concurrency conflict (or the row's since been deleted): the row's current
                    // database values no longer match what was loaded. Keep checking the rest of
                    // the batch so every conflict from one Save click is reported together.
                    conflictedRowKeys.Add(primaryKeys.ToDictionary(key => key, key => update.OriginalKeyValues[key]));
                    continue;
                }

                affectedRows += rowsAffected;
            }

            if (conflictedRowKeys.Count > 0)
            {
                await transaction.RollbackAsync(cancellationToken);
                return new RowSaveResult
                {
                    TotalIntendedChanges = totalIntendedChanges,
                    AffectedRows = 0,
                    WasExecuted = false,
                    ConflictedRowKeys = conflictedRowKeys
                };
            }

            foreach (var insert in rowInserts)
            {
                var valueColumns = insertableColumns
                    .Where(name => insert.Values.ContainsKey(name))
                    .ToList();

                if (valueColumns.Count == 0)
                {
                    continue;
                }

                var sql = RowEditSqlBuilder.BuildInsertStatement(schemaName, tableName, valueColumns);

                await using var command = new SqlCommand(sql, connection, transaction)
                {
                    CommandTimeout = commandTimeoutSeconds
                };

                foreach (var column in valueColumns)
                {
                    command.Parameters.AddWithValue($"@ins_{column}", ToDbValue(insert.Values[column]));
                }

                affectedRows += await command.ExecuteNonQueryAsync(cancellationToken);
            }

            await transaction.CommitAsync(cancellationToken);
            return new RowSaveResult
            {
                TotalIntendedChanges = totalIntendedChanges,
                AffectedRows = affectedRows,
                WasExecuted = true,
                ConflictedRowKeys = []
            };
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

        var sql = RowEditSqlBuilder.BuildDeleteByPrimaryKeyStatement(schemaName, tableName, primaryKeys);

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

    public async Task<RowDeleteByColumnsResult> DeleteRowsBySelectedColumnsAsync(
        string connectionString,
        string schemaName,
        string tableName,
        IReadOnlyList<string> selectedColumns,
        IReadOnlyList<IReadOnlyDictionary<string, object?>> selectedRows,
        int commandTimeoutSeconds,
        CancellationToken cancellationToken)
    {
        if (selectedColumns.Count == 0)
        {
            throw new InvalidOperationException("At least one column must be selected to delete rows without a primary key.");
        }

        if (selectedRows.Count == 0)
        {
            return new RowDeleteByColumnsResult
            {
                IntendedRows = 0,
                MatchedRows = 0,
                DeletedRows = 0,
                WasExecuted = false,
                GeneratedDeleteSql = string.Empty
            };
        }

        var generatedSql = RowEditSqlBuilder.BuildGeneratedDeleteSql(schemaName, tableName, selectedColumns, selectedRows);

        await using var connection = new SqlConnection(connectionString);
        await connection.OpenAsync(cancellationToken);
        await using var transaction = (SqlTransaction)await connection.BeginTransactionAsync(cancellationToken);

        try
        {
            var matchedRows = 0;

            foreach (var row in selectedRows)
            {
                var countSql = RowEditSqlBuilder.BuildCountSql(schemaName, tableName, selectedColumns);
                await using var countCommand = new SqlCommand(countSql, connection, transaction)
                {
                    CommandTimeout = commandTimeoutSeconds
                };

                AddPredicateParameters(countCommand, selectedColumns, row, string.Empty);
                var actualCount = Convert.ToInt32(await countCommand.ExecuteScalarAsync(cancellationToken), CultureInfo.InvariantCulture);

                if (actualCount > 0)
                {
                    matchedRows++;
                }
            }

            if (matchedRows != selectedRows.Count)
            {
                await transaction.RollbackAsync(cancellationToken);
                return new RowDeleteByColumnsResult
                {
                    IntendedRows = selectedRows.Count,
                    MatchedRows = matchedRows,
                    DeletedRows = 0,
                    WasExecuted = false,
                    GeneratedDeleteSql = generatedSql
                };
            }

            var deletedRows = 0;
            foreach (var row in selectedRows)
            {
                var deleteSql = RowEditSqlBuilder.BuildDeleteTopOneSql(schemaName, tableName, selectedColumns);
                await using var deleteCommand = new SqlCommand(deleteSql, connection, transaction)
                {
                    CommandTimeout = commandTimeoutSeconds
                };

                AddPredicateParameters(deleteCommand, selectedColumns, row, string.Empty);
                deletedRows += await deleteCommand.ExecuteNonQueryAsync(cancellationToken);
            }

            await transaction.CommitAsync(cancellationToken);

            return new RowDeleteByColumnsResult
            {
                IntendedRows = selectedRows.Count,
                MatchedRows = matchedRows,
                DeletedRows = deletedRows,
                WasExecuted = true,
                GeneratedDeleteSql = generatedSql
            };
        }
        catch
        {
            await transaction.RollbackAsync(cancellationToken);
            throw;
        }
    }

    private static object ToDbValue(object? value)
    {
        return value ?? DBNull.Value;
    }

    private static void EnsureMatchValues(IReadOnlyList<string> matchColumns, IReadOnlyDictionary<string, object?> keyValues)
    {
        foreach (var key in matchColumns)
        {
            if (!keyValues.ContainsKey(key))
            {
                throw new InvalidOperationException($"Original value is required for column '{key}'.");
            }
        }
    }

    private static void EnsurePrimaryKeyValues(IReadOnlyList<string> primaryKeys, IReadOnlyDictionary<string, object?> keyValues)
    {
        EnsureMatchValues(primaryKeys, keyValues);
    }

    private static void AddPredicateParameters(
        SqlCommand command,
        IReadOnlyList<string> selectedColumns,
        IReadOnlyDictionary<string, object?> row,
        string parameterPrefix)
    {
        foreach (var column in selectedColumns)
        {
            if (!row.TryGetValue(column, out var value))
            {
                throw new InvalidOperationException($"Selected row does not include value for column '{column}'.");
            }

            command.Parameters.AddWithValue($"@{parameterPrefix}{column}", ToDbValue(value));
        }
    }
}
