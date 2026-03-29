namespace DatabaseManager.Core.Models.Schema;

public sealed class StoredProcedureParameterInfo
{
    public required string ParameterName { get; init; }

    public required string DataType { get; init; }

    public int OrdinalPosition { get; init; }

    public bool IsOutput { get; init; }

    public bool HasDefaultValue { get; init; }

    public short MaxLength { get; init; }

    public byte Precision { get; init; }

    public byte Scale { get; init; }

    public bool IsReturnValue => OrdinalPosition == 0;
}
