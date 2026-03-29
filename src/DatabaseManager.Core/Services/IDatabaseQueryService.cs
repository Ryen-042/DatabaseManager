using DatabaseManager.Core.Models;

namespace DatabaseManager.Core.Services;

public interface IDatabaseQueryService
{
    Task<QueryExecutionResult> ExecuteAsync(
        string connectionString,
        string sql,
        int commandTimeoutSeconds,
        CancellationToken cancellationToken);
}
