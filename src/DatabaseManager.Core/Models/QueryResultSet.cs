using System.Data;

namespace DatabaseManager.Core.Models;

public sealed class QueryResultSet
{
    public required string Title { get; init; }

    public DataTable? DataTable { get; init; }

    public int AffectedRows { get; init; }
}
