using System.Data;

namespace DatabaseManager.Core.Models;

public sealed class QueryExecutionResult
{
    public required bool IsSuccess { get; init; }

    public string? ErrorMessage { get; init; }

    public DataTable? DataTable { get; init; }

    public int AffectedRows { get; init; }

    public TimeSpan Duration { get; init; }

    public IReadOnlyList<QueryResultSet>? ResultSets { get; init; }

    public IReadOnlyDictionary<string, object?>? OutputParameters { get; init; }
}
