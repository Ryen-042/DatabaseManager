namespace DatabaseManager.Core.Models.Editing;

public sealed class RowDeleteByColumnsResult
{
    public int IntendedRows { get; init; }

    public int MatchedRows { get; init; }

    public int DeletedRows { get; init; }

    public bool WasExecuted { get; init; }

    public required string GeneratedDeleteSql { get; init; }
}
