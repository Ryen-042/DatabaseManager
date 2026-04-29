using System.Data;
using System.Diagnostics;
using DatabaseManager.Core.Models;
using Microsoft.Data.SqlClient;

namespace DatabaseManager.Core.Services;

public sealed class SqlServerQueryService : IDatabaseQueryService
{
    public async Task<QueryExecutionResult> ExecuteAsync(
        string connectionString,
        string sql,
        int commandTimeoutSeconds,
        CancellationToken cancellationToken)
    {
        return await ExecuteInternalAsync(
            connectionString,
            sql,
            parameters: null,
            commandTimeoutSeconds,
            cancellationToken);
    }

    public async Task<QueryExecutionResult> ExecuteAsync(
        string connectionString,
        string sql,
        IReadOnlyList<QueryParameterValue> parameters,
        int commandTimeoutSeconds,
        CancellationToken cancellationToken)
    {
        return await ExecuteInternalAsync(
            connectionString,
            sql,
            parameters,
            commandTimeoutSeconds,
            cancellationToken);
    }

    private static async Task<QueryExecutionResult> ExecuteInternalAsync(
        string connectionString,
        string sql,
        IReadOnlyList<QueryParameterValue>? parameters,
        int commandTimeoutSeconds,
        CancellationToken cancellationToken)
    {
        var stopwatch = Stopwatch.StartNew();

        try
        {
            await using var connection = new SqlConnection(connectionString);
            await connection.OpenAsync(cancellationToken);

            await using var command = new SqlCommand(sql, connection)
            {
                CommandTimeout = commandTimeoutSeconds
            };

            if (parameters is not null)
            {
                foreach (var parameter in parameters)
                {
                    command.Parameters.AddWithValue(parameter.Name, parameter.Value ?? DBNull.Value);
                }
            }

            var resultSets = new List<QueryResultSet>();
            await using var reader = await command.ExecuteReaderAsync(cancellationToken);

            var resultSetIndex = 1;
            do
            {
                if (reader.FieldCount <= 0)
                {
                    continue;
                }

                var dataTable = await ReadCurrentResultSetAsync(reader, cancellationToken);

                resultSets.Add(new QueryResultSet
                {
                    Title = $"Result Set {resultSetIndex++}",
                    DataTable = dataTable,
                    AffectedRows = dataTable.Rows.Count
                });
            }
            while (await reader.NextResultAsync(cancellationToken));

            var affectedRows = reader.RecordsAffected;
            if (resultSets.Count == 0)
            {
                resultSets.Add(new QueryResultSet
                {
                    Title = "Statement Summary",
                    DataTable = null,
                    AffectedRows = Math.Max(0, affectedRows)
                });
            }
            else if (affectedRows > 0)
            {
                resultSets.Add(new QueryResultSet
                {
                    Title = "Statement Summary",
                    DataTable = null,
                    AffectedRows = affectedRows
                });
            }

            var primaryDataTable = resultSets.FirstOrDefault(x => x.DataTable is not null)?.DataTable;
            stopwatch.Stop();

            return new QueryExecutionResult
            {
                IsSuccess = true,
                DataTable = primaryDataTable,
                AffectedRows = primaryDataTable is null ? Math.Max(0, affectedRows) : primaryDataTable.Rows.Count,
                Duration = stopwatch.Elapsed,
                ResultSets = resultSets
            };
        }
        catch (Exception ex)
        {
            stopwatch.Stop();
            return new QueryExecutionResult
            {
                IsSuccess = false,
                ErrorMessage = ex.Message,
                DataTable = null,
                AffectedRows = 0,
                Duration = stopwatch.Elapsed,
                ResultSets = null
            };
        }
    }

    private static async Task<DataTable> ReadCurrentResultSetAsync(SqlDataReader reader, CancellationToken cancellationToken)
    {
        var table = new DataTable();

        for (var i = 0; i < reader.FieldCount; i++)
        {
            var baseColumnName = reader.GetName(i);
            var columnName = GetUniqueColumnName(table, string.IsNullOrWhiteSpace(baseColumnName) ? $"Column{i + 1}" : baseColumnName);
            table.Columns.Add(columnName, reader.GetFieldType(i));
        }

        while (await reader.ReadAsync(cancellationToken))
        {
            var values = new object[reader.FieldCount];
            reader.GetValues(values);
            table.Rows.Add(values);
        }

        return table;
    }

    private static string GetUniqueColumnName(DataTable table, string baseName)
    {
        if (!table.Columns.Contains(baseName))
        {
            return baseName;
        }

        var suffix = 1;
        var candidate = $"{baseName}_{suffix}";
        while (table.Columns.Contains(candidate))
        {
            suffix++;
            candidate = $"{baseName}_{suffix}";
        }

        return candidate;
    }
}
