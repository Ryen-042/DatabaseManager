using DatabaseManager.Core.Models;
using DatabaseManager.Wpf.ViewModels;

namespace DatabaseManager.Tests.ViewModels;

public sealed class ConnectionProfileViewModelTests
{
    private static ConnectionProfile CreateProfile(bool isDefault = false, DateTimeOffset? lastUsedAtUtc = null) => new()
    {
        Id = Guid.NewGuid(),
        Name = "Sales DB",
        EncryptedConnectionString = "irrelevant-ciphertext",
        IsDefault = isDefault,
        CreatedAtUtc = DateTimeOffset.UtcNow,
        UpdatedAtUtc = DateTimeOffset.UtcNow,
        LastUsedAtUtc = lastUsedAtUtc
    };

    [Fact]
    public void Properties_ReflectTheUnderlyingProfileAndConstructorPreview()
    {
        var profile = CreateProfile(isDefault: true);

        var vm = new ConnectionProfileViewModel(profile, "Server=myserver;Initial Catalog=Sales;");

        Assert.Equal(profile.Id, vm.Id);
        Assert.Equal("Sales DB", vm.Name);
        Assert.True(vm.IsDefault);
        Assert.Equal("Server=myserver;Initial Catalog=Sales;", vm.ConnectionStringPreview);
        Assert.Same(profile, vm.Profile);
    }

    [Fact]
    public void LastUsedDisplay_NoLastUsed_ReadsNeverUsed()
    {
        var vm = new ConnectionProfileViewModel(CreateProfile(), "Server=myserver;");

        Assert.Equal("Never used", vm.LastUsedDisplay);
    }

    [Fact]
    public void LastUsedDisplay_HasLastUsed_FormatsIt()
    {
        var lastUsed = new DateTimeOffset(2026, 1, 15, 9, 30, 0, TimeSpan.Zero);
        var vm = new ConnectionProfileViewModel(CreateProfile(lastUsedAtUtc: lastUsed), "Server=myserver;");

        Assert.NotEqual("Never used", vm.LastUsedDisplay);
    }
}
