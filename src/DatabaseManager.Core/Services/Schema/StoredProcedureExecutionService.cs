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
            var dataTable = new DataTable();
            dataTable.Load(reader);

            var outputValues = new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase);
            foreach (SqlParameter sqlParameter in command.Parameters)
            {
                if (sqlParameter.Direction is ParameterDirection.Output or ParameterDirection.InputOutput)
                {
                    outputValues[sqlParameter.ParameterName] = sqlParameter.Value == DBNull.Value ? null : sqlParameter.Value;
                }
            }

            stopwatch.Stop();
            return new QueryExecutionResult
            {
                IsSuccess = true,
                DataTable = dataTable.Rows.Count > 0 ? dataTable : null,
                AffectedRows = dataTable.Rows.Count,
                Duration = stopwatch.Elapsed,
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
                Duration = stopwatch.Elapsed
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
}
