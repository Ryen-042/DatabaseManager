using DatabaseManager.Core.Models;

namespace DatabaseManager.Wpf.ViewModels;

/// <summary>
/// Thin bindable wrapper around a <see cref="ConnectionProfile"/> for display in the connection
/// picker's list. Takes the already-decrypted connection string at construction (the caller
/// already needs <c>IConnectionProfileStoreService.Decrypt</c> for other purposes) rather than
/// holding the store itself, so this stays a plain display wrapper with no service dependency.
/// </summary>
public sealed class ConnectionProfileViewModel(ConnectionProfile profile, string connectionStringPreview)
{
    public ConnectionProfile Profile { get; } = profile;

    public Guid Id => Profile.Id;

    public string Name => Profile.Name;

    public bool IsDefault => Profile.IsDefault;

    /// <summary>The decrypted connection string, shown dimmed/trimmed under the name in the picker's list.</summary>
    public string ConnectionStringPreview { get; } = connectionStringPreview;

    public string LastUsedDisplay => Profile.LastUsedAtUtc.HasValue
        ? Profile.LastUsedAtUtc.Value.LocalDateTime.ToString("g")
        : "Never used";
}
