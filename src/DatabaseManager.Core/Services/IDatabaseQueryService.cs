using DatabaseManager.Core.Models;

namespace DatabaseManager.Core.Services;

public interface IDatabaseQueryService
{
    Task<QueryExecutionResult> ExecuteAsync(
        string connectionString,
        string sql,
        int commandTimeoutSeconds,
        CancellationToken cancellationToken);

    Task<QueryExecutionResult> ExecuteAsync(
        string connectionString,
        string sql,
        IReadOnlyList<QueryParameterValue> parameters,
        int commandTimeoutSeconds,
        CancellationToken cancellationToken);
}
