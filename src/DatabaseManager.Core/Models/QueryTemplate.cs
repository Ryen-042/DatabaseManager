namespace DatabaseManager.Core.Models;

public sealed class QueryTemplate
{
    public required string Name { get; init; }

    public required string Sql { get; init; }

    public DateTimeOffset CreatedAtUtc { get; init; } = DateTimeOffset.UtcNow;

    public DateTimeOffset UpdatedAtUtc { get; init; } = DateTimeOffset.UtcNow;
}
