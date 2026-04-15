namespace DatabaseManager.Core.Models.Editing;

public sealed class RowInsertRequest
{
    public required IReadOnlyDictionary<string, object?> Values { get; init; }
}
