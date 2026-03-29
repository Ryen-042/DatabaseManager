namespace DatabaseManager.Core.Models.Editing;

public sealed class RowUpdateRequest
{
    public required IReadOnlyDictionary<string, object?> OriginalKeyValues { get; init; }

    public required IReadOnlyDictionary<string, object?> CurrentValues { get; init; }
}
