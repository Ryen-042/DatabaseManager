using DatabaseManager.Core.Models;
using DatabaseManager.Core.Models.Schema;

namespace DatabaseManager.Core.Services.Schema;

public interface IStoredProcedureExecutionService
{
    Task<QueryExecutionResult> ExecuteAsync(
        string connectionString,
        string schemaName,
        string procedureName,
        IReadOnlyList<StoredProcedureExecutionParameter> parameters,
        int commandTimeoutSeconds,
        CancellationToken cancellationToken);
}
