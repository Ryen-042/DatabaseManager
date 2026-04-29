using System.Data;
using System.Diagnostics;
using DatabaseManager.Core.Models;
using DatabaseManager.Core.Models.Schema;
using Microsoft.Data.SqlClient;

namespace DatabaseManager.Core.Services.Schema;

public sealed class StoredProcedureExecutionService : IStoredProcedureExecutionService
{
    public async Task<QueryExecutionResult> ExecuteAsync(
        string connectionString,
        string schemaName,
        string procedureName,
        IReadOnlyList<StoredProcedureExecutionParameter> parameters,
        int commandTimeoutSeconds,
        CancellationToken cancellationToken)
    {
        var stopwatch = Stopwatch.StartNew();

        try
        {
            await using var connection = new SqlConnection(connectionString);
            await connection.OpenAsync(cancellationToken);

            await using var command = new SqlCommand($"[{schemaName}].[{procedureName}]", connection)
            {
                CommandType = CommandType.StoredProcedure,
                CommandTimeout = commandTimeoutSeconds
            };

            foreach (var parameter in parameters)
            {
                var sqlParameter = new SqlParameter(parameter.Name, GetInputValue(parameter));
                if (parameter.IsInputOutput)
                {
                    sqlParameter.Direction = ParameterDirection.InputOutput;
                }
                else if (parameter.IsOutput)
                {
                    sqlParameter.Direction = ParameterDirection.Output;
                }

                command.Parameters.Add(sqlParameter);
            }

            await using var reader = await command.ExecuteReaderAsync(cancellationToken);
            var resultSets = new List<QueryResultSet>();
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

            if (resultSets.Count == 0)
            {
                resultSets.Add(new QueryResultSet
                {
                    Title = "Statement Summary",
                    DataTable = null,
                    AffectedRows = Math.Max(0, reader.RecordsAffected)
                });
            }
            else if (reader.RecordsAffected > 0)
            {
                resultSets.Add(new QueryResultSet
                {
                    Title = "Statement Summary",
                    DataTable = null,
                    AffectedRows = reader.RecordsAffected
                });
            }

            var outputValues = new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase);
            foreach (SqlParameter sqlParameter in command.Parameters)
            {
                if (sqlParameter.Direction is ParameterDirection.Output or ParameterDirection.InputOutput)
                {
                    outputValues[sqlParameter.ParameterName] = sqlParameter.Value == DBNull.Value ? null : sqlParameter.Value;
                }
            }

            stopwatch.Stop();
            var primaryDataTable = resultSets.FirstOrDefault(x => x.DataTable is not null)?.DataTable;
            return new QueryExecutionResult
            {
                IsSuccess = true,
                DataTable = primaryDataTable,
                AffectedRows = primaryDataTable?.Rows.Count ?? Math.Max(0, reader.RecordsAffected),
                Duration = stopwatch.Elapsed,
                ResultSets = resultSets,
                OutputParameters = outputValues.Count == 0 ? null : outputValues
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

    private static object GetInputValue(StoredProcedureExecutionParameter parameter)
    {
        if (parameter.SendAsNull)
        {
            return DBNull.Value;
        }

        if (string.IsNullOrWhiteSpace(parameter.Value))
        {
            return DBNull.Value;
        }

        return parameter.Value;
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
