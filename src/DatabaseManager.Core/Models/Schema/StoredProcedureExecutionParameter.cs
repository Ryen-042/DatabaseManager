namespace DatabaseManager.Core.Models.Schema;

public sealed class StoredProcedureExecutionParameter
{
    public required string Name { get; init; }

    public string? Value { get; init; }

    public bool IsOutput { get; init; }

    public bool IsInputOutput { get; init; }

    public bool SendAsNull { get; init; }
}
