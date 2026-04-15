namespace DatabaseManager.Core.Models;

public sealed class QueryParameterValue
{
    public required string Name { get; init; }

    public object? Value { get; init; }
}
