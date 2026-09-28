namespace DatabaseManager.Core.Models.Editing;

/// <summary>
/// Result of RowEditService.SaveRowChangesAsync. WasExecuted is false when one or more
/// primary-keyed updates hit a concurrency conflict (its WHERE clause matched 0 rows because the
/// row's original values no longer match what's actually in the database) - in that case nothing
/// was committed (all-or-nothing, same as the no-PK delete mismatch safeguard) and
/// ConflictedRowKeys lists every conflicting row's primary-key values so the caller can report
/// exactly which ones need a reload.
/// </summary>
public sealed class RowSaveResult
{
    public required int TotalIntendedChanges { get; init; }

    public int AffectedRows { get; init; }

    public bool WasExecuted { get; init; }

    public required IReadOnlyList<IReadOnlyDictionary<string, object?>> ConflictedRowKeys { get; init; }
}
