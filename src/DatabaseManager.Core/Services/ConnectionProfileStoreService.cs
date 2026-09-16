using System.Runtime.Versioning;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using DatabaseManager.Core.Models;

namespace DatabaseManager.Core.Services;

/// <summary>
/// JSON-persisted connection profile store, mirroring <see cref="TemplateStoreService"/>'s
/// pattern (same %LocalAppData%/DatabaseManager location, same missing-file tolerance).
/// Connection strings are DPAPI-encrypted (current-user scope) before they ever reach disk.
/// </summary>
[SupportedOSPlatform("windows")]
public sealed class ConnectionProfileStoreService : IConnectionProfileStoreService
{
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        WriteIndented = true,
        PropertyNameCaseInsensitive = true
    };

    private readonly string _filePath;

    public ConnectionProfileStoreService(string? filePath = null)
    {
        var appData = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "DatabaseManager");

        Directory.CreateDirectory(appData);
        _filePath = filePath ?? Path.Combine(appData, "connection-profiles.json");
    }

    public async Task<IReadOnlyList<ConnectionProfile>> GetAllAsync(CancellationToken cancellationToken)
    {
        if (!File.Exists(_filePath))
        {
            return Array.Empty<ConnectionProfile>();
        }

        await using var stream = File.OpenRead(_filePath);
        var items = await JsonSerializer.DeserializeAsync<List<ConnectionProfile>>(stream, JsonOptions, cancellationToken);

        if (items is null)
        {
            return Array.Empty<ConnectionProfile>();
        }

        return items
            .OrderBy(x => x.Name, StringComparer.OrdinalIgnoreCase)
            .ToList();
    }

    public async Task<ConnectionProfile> SaveAsync(
        Guid? id,
        string name,
        string plainConnectionString,
        bool isDefault,
        CancellationToken cancellationToken)
    {
        if (string.IsNullOrWhiteSpace(name))
        {
            throw new ArgumentException("Profile name is required.", nameof(name));
        }

        if (string.IsNullOrWhiteSpace(plainConnectionString))
        {
            throw new ArgumentException("Connection string is required.", nameof(plainConnectionString));
        }

        var profiles = (await GetAllAsync(cancellationToken)).ToList();
        var now = DateTimeOffset.UtcNow;
        var encrypted = Encrypt(plainConnectionString);

        var existingIndex = id.HasValue ? profiles.FindIndex(p => p.Id == id.Value) : -1;

        ConnectionProfile saved = existingIndex >= 0
            ? profiles[existingIndex] with
            {
                Name = name.Trim(),
                EncryptedConnectionString = encrypted,
                IsDefault = isDefault,
                UpdatedAtUtc = now
            }
            : new ConnectionProfile
            {
                Id = Guid.NewGuid(),
                Name = name.Trim(),
                EncryptedConnectionString = encrypted,
                IsDefault = isDefault,
                CreatedAtUtc = now,
                UpdatedAtUtc = now
            };

        if (existingIndex >= 0)
        {
            profiles[existingIndex] = saved;
        }
        else
        {
            profiles.Add(saved);
        }

        if (isDefault)
        {
            for (var i = 0; i < profiles.Count; i++)
            {
                if (profiles[i].Id != saved.Id && profiles[i].IsDefault)
                {
                    profiles[i] = profiles[i] with { IsDefault = false };
                }
            }
        }

        await WriteAllAsync(profiles, cancellationToken);
        return saved;
    }

    public async Task DeleteAsync(Guid id, CancellationToken cancellationToken)
    {
        var profiles = (await GetAllAsync(cancellationToken)).ToList();
        profiles.RemoveAll(p => p.Id == id);
        await WriteAllAsync(profiles, cancellationToken);
    }

    public async Task TouchLastUsedAsync(Guid id, CancellationToken cancellationToken)
    {
        var profiles = (await GetAllAsync(cancellationToken)).ToList();
        var index = profiles.FindIndex(p => p.Id == id);
        if (index < 0)
        {
            return;
        }

        profiles[index] = profiles[index] with { LastUsedAtUtc = DateTimeOffset.UtcNow };
        await WriteAllAsync(profiles, cancellationToken);
    }

    public string Decrypt(ConnectionProfile profile) => Decrypt(profile.EncryptedConnectionString);

    private static string Encrypt(string plainText)
    {
        var plainBytes = Encoding.UTF8.GetBytes(plainText);
        var encryptedBytes = ProtectedData.Protect(plainBytes, optionalEntropy: null, DataProtectionScope.CurrentUser);
        return Convert.ToBase64String(encryptedBytes);
    }

    private static string Decrypt(string cipherTextBase64)
    {
        var encryptedBytes = Convert.FromBase64String(cipherTextBase64);
        var plainBytes = ProtectedData.Unprotect(encryptedBytes, optionalEntropy: null, DataProtectionScope.CurrentUser);
        return Encoding.UTF8.GetString(plainBytes);
    }

    private async Task WriteAllAsync(IReadOnlyCollection<ConnectionProfile> profiles, CancellationToken cancellationToken)
    {
        await using var stream = File.Create(_filePath);
        await JsonSerializer.SerializeAsync(stream, profiles, JsonOptions, cancellationToken);
    }
}
