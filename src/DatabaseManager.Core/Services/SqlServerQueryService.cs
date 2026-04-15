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

            if (LooksLikeReadQuery(sql))
            {
                await using var reader = await command.ExecuteReaderAsync(cancellationToken);
                var dataTable = new DataTable();
                dataTable.Load(reader);

                stopwatch.Stop();
                return new QueryExecutionResult
                {
                    IsSuccess = true,
                    DataTable = dataTable,
                    AffectedRows = dataTable.Rows.Count,
                    Duration = stopwatch.Elapsed
                };
            }

            var affectedRows = await command.ExecuteNonQueryAsync(cancellationToken);
            stopwatch.Stop();

            return new QueryExecutionResult
            {
                IsSuccess = true,
                DataTable = null,
                AffectedRows = affectedRows,
                Duration = stopwatch.Elapsed
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
                Duration = stopwatch.Elapsed
            };
        }
    }

    private static bool LooksLikeReadQuery(string sql)
    {
        var trimmed = sql.TrimStart();
        return trimmed.StartsWith("SELECT", StringComparison.OrdinalIgnoreCase)
            || trimmed.StartsWith("WITH", StringComparison.OrdinalIgnoreCase)
            || trimmed.StartsWith("EXEC", StringComparison.OrdinalIgnoreCase)
            || trimmed.StartsWith("SHOW", StringComparison.OrdinalIgnoreCase);
    }
}
