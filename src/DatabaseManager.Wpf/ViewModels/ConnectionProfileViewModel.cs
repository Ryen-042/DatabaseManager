using DatabaseManager.Core.Models;

namespace DatabaseManager.Wpf.ViewModels;

/// <summary>Thin bindable wrapper around a <see cref="ConnectionProfile"/> for display in the connection picker's list.</summary>
public sealed class ConnectionProfileViewModel(ConnectionProfile profile)
{
    public ConnectionProfile Profile { get; } = profile;

    public Guid Id => Profile.Id;

    public string Name => Profile.Name;

    public bool IsDefault => Profile.IsDefault;

    public string LastUsedDisplay => Profile.LastUsedAtUtc.HasValue
        ? Profile.LastUsedAtUtc.Value.LocalDateTime.ToString("g")
        : "Never used";
}
