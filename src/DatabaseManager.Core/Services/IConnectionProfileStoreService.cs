using DatabaseManager.Core.Models;

namespace DatabaseManager.Core.Services;

public interface IConnectionProfileStoreService
{
    Task<IReadOnlyList<ConnectionProfile>> GetAllAsync(CancellationToken cancellationToken);

    /// <summary>
    /// Creates a new profile (<paramref name="id"/> is null) or updates an existing one.
    /// <paramref name="plainConnectionString"/> is encrypted before being persisted.
    /// </summary>
    Task<ConnectionProfile> SaveAsync(
        Guid? id,
        string name,
        string plainConnectionString,
        bool isDefault,
        CancellationToken cancellationToken);

    Task DeleteAsync(Guid id, CancellationToken cancellationToken);

    Task TouchLastUsedAsync(Guid id, CancellationToken cancellationToken);

    /// <summary>Decrypts a profile's connection string. Never persisted - in-memory only.</summary>
    string Decrypt(ConnectionProfile profile);
}
