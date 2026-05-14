using System.Data;
using System.Globalization;
using System.Text;
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
        var sql = $"SELECT TOP (@topRows) * FROM [{EscapeIdentifier(schemaName)}].[{EscapeIdentifier(tableName)}]";
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

    public async Task<int> SaveUpdatedRowsAsync(
        string connectionString,
        string schemaName,
        string tableName,
        IReadOnlyList<ColumnSchemaInfo> columns,
        IReadOnlyList<RowUpdateRequest> rowUpdates,
        int commandTimeoutSeconds,
        CancellationToken cancellationToken)
    {
        return await SaveRowChangesAsync(
            connectionString,
            schemaName,
            tableName,
            columns,
            rowUpdates,
            Array.Empty<RowInsertRequest>(),
            commandTimeoutSeconds,
            cancellationToken);
    }

    public async Task<int> SaveRowChangesAsync(
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
        var matchColumns = primaryKeys.Count > 0
            ? primaryKeys
            : columns.Select(c => c.ColumnName).ToList();

        var updatableColumns = columns
            .Where(c => !c.IsPrimaryKey && !c.IsIdentity && !IsServerGeneratedColumn(c))
            .Select(c => c.ColumnName)
            .ToList();

        var insertableColumns = columns
            .Where(c => !c.IsIdentity && !IsServerGeneratedColumn(c))
            .Select(c => c.ColumnName)
            .ToList();

        if ((updatableColumns.Count == 0 || rowUpdates.Count == 0) && rowInserts.Count == 0)
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

                EnsureMatchValues(matchColumns, update.OriginalKeyValues);

                var setClause = string.Join(", ", setColumns.Select(c => $"[{EscapeIdentifier(c)}] = @set_{c}"));
                var whereClause = BuildNullSafeWhereClause(matchColumns);
                var sql = $"UPDATE [{EscapeIdentifier(schemaName)}].[{EscapeIdentifier(tableName)}] SET {setClause} WHERE {whereClause};";

                if (primaryKeys.Count == 0)
                {
                    sql += Environment.NewLine
                        + "IF @@ROWCOUNT <> 1" + Environment.NewLine
                        + "    THROW 50000, 'Cannot save changes because the row is not uniquely identifiable without a primary key.', 1;";
                }

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

                affectedRows += await command.ExecuteNonQueryAsync(cancellationToken);
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

                var columnClause = string.Join(", ", valueColumns.Select(c => $"[{EscapeIdentifier(c)}]"));
                var valuesClause = string.Join(", ", valueColumns.Select(c => $"@ins_{c}"));
                var sql = $"INSERT INTO [{EscapeIdentifier(schemaName)}].[{EscapeIdentifier(tableName)}] ({columnClause}) VALUES ({valuesClause});";

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

        var generatedSql = BuildGeneratedDeleteSql(schemaName, tableName, selectedColumns, selectedRows);

        await using var connection = new SqlConnection(connectionString);
        await connection.OpenAsync(cancellationToken);
        await using var transaction = (SqlTransaction)await connection.BeginTransactionAsync(cancellationToken);

        try
        {
            var matchedRows = 0;

            foreach (var row in selectedRows)
            {
                var countSql = BuildCountSql(schemaName, tableName, selectedColumns);
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
                var deleteSql = BuildDeleteTopOneSql(schemaName, tableName, selectedColumns);
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

    private static string EscapeIdentifier(string value)
    {
        return value.Replace("]", "]]", StringComparison.Ordinal);
    }

    private static object ToDbValue(object? value)
    {
        return value ?? DBNull.Value;
    }

    private static bool IsServerGeneratedColumn(ColumnSchemaInfo column)
    {
        return column.DataType.Equals("rowversion", StringComparison.OrdinalIgnoreCase)
            || column.DataType.Equals("timestamp", StringComparison.OrdinalIgnoreCase);
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

    private static string BuildNullSafeWhereClause(IReadOnlyList<string> columns)
    {
        return string.Join(" AND ", columns.Select(column =>
            $"((@key_{column} IS NULL AND [{EscapeIdentifier(column)}] IS NULL) OR [{EscapeIdentifier(column)}] = @key_{column})"));
    }

    private static string BuildCountSql(string schemaName, string tableName, IReadOnlyList<string> selectedColumns)
    {
        var whereClause = BuildWhereClauseForSelectedColumns(selectedColumns, parameterPrefix: string.Empty);
        return $"SELECT COUNT(1) FROM [{EscapeIdentifier(schemaName)}].[{EscapeIdentifier(tableName)}] WHERE {whereClause};";
    }

    private static string BuildDeleteTopOneSql(string schemaName, string tableName, IReadOnlyList<string> selectedColumns)
    {
        var whereClause = BuildWhereClauseForSelectedColumns(selectedColumns, parameterPrefix: string.Empty);
        return $"DELETE TOP (1) FROM [{EscapeIdentifier(schemaName)}].[{EscapeIdentifier(tableName)}] WHERE {whereClause};";
    }

    private static string BuildWhereClauseForSelectedColumns(IReadOnlyList<string> selectedColumns, string parameterPrefix)
    {
        return string.Join(" AND ", selectedColumns.Select(column =>
            $"((@{parameterPrefix}{column} IS NULL AND [{EscapeIdentifier(column)}] IS NULL) OR [{EscapeIdentifier(column)}] = @{parameterPrefix}{column})"));
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

    private static string BuildGeneratedDeleteSql(
        string schemaName,
        string tableName,
        IReadOnlyList<string> selectedColumns,
        IReadOnlyList<IReadOnlyDictionary<string, object?>> selectedRows)
    {
        var sb = new StringBuilder();
        sb.AppendLine($"-- Generated delete script for [{EscapeIdentifier(schemaName)}].[{EscapeIdentifier(tableName)}]");
        sb.AppendLine("-- Review before execution.");
        sb.AppendLine();

        for (var i = 0; i < selectedRows.Count; i++)
        {
            var row = selectedRows[i];
            var predicates = new List<string>();

            foreach (var column in selectedColumns)
            {
                if (!row.TryGetValue(column, out var value))
                {
                    throw new InvalidOperationException($"Selected row does not include value for column '{column}'.");
                }

                if (value is null or DBNull)
                {
                    predicates.Add($"[{EscapeIdentifier(column)}] IS NULL");
                }
                else
                {
                    predicates.Add($"[{EscapeIdentifier(column)}] = {ToSqlLiteral(value)}");
                }
            }

            sb.AppendLine($"-- Row {i + 1}");
            sb.AppendLine($"DELETE TOP (1) FROM [{EscapeIdentifier(schemaName)}].[{EscapeIdentifier(tableName)}]");
            sb.AppendLine($"WHERE {string.Join(" AND ", predicates)};");
            sb.AppendLine("GO");
            sb.AppendLine();
        }

        return sb.ToString().TrimEnd();
    }

    private static string ToSqlLiteral(object value)
    {
        return value switch
        {
            string s => $"N'{s.Replace("'", "''", StringComparison.Ordinal)}'",
            char c => $"N'{c.ToString().Replace("'", "''", StringComparison.Ordinal)}'",
            bool b => b ? "1" : "0",
            DateTime dt => $"'{dt:yyyy-MM-dd HH:mm:ss.fffffff}'",
            DateTimeOffset dto => $"'{dto:yyyy-MM-dd HH:mm:ss.fffffff zzz}'",
            byte[] bytes => $"0x{Convert.ToHexString(bytes)}",
            Guid guid => $"'{guid:D}'",
            IFormattable formattable => formattable.ToString(null, CultureInfo.InvariantCulture) ?? "NULL",
            _ => $"N'{value.ToString()?.Replace("'", "''", StringComparison.Ordinal)}'"
        };
    }
}
