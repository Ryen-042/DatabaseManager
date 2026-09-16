namespace DatabaseManager.Core.Models;

/// <summary>
/// A saved connection profile. <see cref="EncryptedConnectionString"/> is DPAPI-protected
/// (current-user scope) at rest via <c>ConnectionProfileStoreService</c> - the plaintext
/// connection string (which may carry a SQL-auth password) is only ever held in memory,
/// never written to disk.
/// </summary>
public sealed record ConnectionProfile
{
    public required Guid Id { get; init; }

    public required string Name { get; init; }

    public required string EncryptedConnectionString { get; init; }

    public bool IsDefault { get; init; }

    public required DateTimeOffset CreatedAtUtc { get; init; }

    public required DateTimeOffset UpdatedAtUtc { get; init; }

    public DateTimeOffset? LastUsedAtUtc { get; init; }
}
