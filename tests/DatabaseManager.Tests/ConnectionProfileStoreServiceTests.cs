using DatabaseManager.Core.Services;

namespace DatabaseManager.Tests;

public sealed class ConnectionProfileStoreServiceTests
{
    private const string SampleConnectionString =
        "Data Source=localhost;Initial Catalog=Test;User ID=sa;Password=Sup3rSecret!;TrustServerCertificate=True;";

    [Fact]
    public async Task GetAllAsync_MissingFile_ReturnsEmpty()
    {
        var tempFile = Path.Combine(Path.GetTempPath(), $"dbm-conn-{Guid.NewGuid():N}.json");
        var service = new ConnectionProfileStoreService(tempFile);

        var items = await service.GetAllAsync(CancellationToken.None);

        Assert.Empty(items);
    }

    [Fact]
    public async Task SaveAsync_ThenGetAllAsync_RoundTripsDecryptedConnectionString()
    {
        var tempFile = Path.Combine(Path.GetTempPath(), $"dbm-conn-{Guid.NewGuid():N}.json");

        try
        {
            var service = new ConnectionProfileStoreService(tempFile);

            var saved = await service.SaveAsync(null, "Local Test", SampleConnectionString, isDefault: false, CancellationToken.None);

            var items = await service.GetAllAsync(CancellationToken.None);

            Assert.Single(items);
            Assert.Equal("Local Test", items[0].Name);
            Assert.Equal(saved.Id, items[0].Id);
            Assert.Equal(SampleConnectionString, service.Decrypt(items[0]));
        }
        finally
        {
            if (File.Exists(tempFile))
            {
                File.Delete(tempFile);
            }
        }
    }

    [Fact]
    public async Task SaveAsync_StoredFileDoesNotContainPlaintextConnectionString()
    {
        var tempFile = Path.Combine(Path.GetTempPath(), $"dbm-conn-{Guid.NewGuid():N}.json");

        try
        {
            var service = new ConnectionProfileStoreService(tempFile);
            await service.SaveAsync(null, "Local Test", SampleConnectionString, isDefault: false, CancellationToken.None);

            var rawFileContents = await File.ReadAllTextAsync(tempFile);

            Assert.DoesNotContain(SampleConnectionString, rawFileContents);
            Assert.DoesNotContain("Sup3rSecret!", rawFileContents);
        }
        finally
        {
            if (File.Exists(tempFile))
            {
                File.Delete(tempFile);
            }
        }
    }

    [Fact]
    public async Task SaveAsync_WithExistingId_UpdatesInPlaceRatherThanAdding()
    {
        var tempFile = Path.Combine(Path.GetTempPath(), $"dbm-conn-{Guid.NewGuid():N}.json");

        try
        {
            var service = new ConnectionProfileStoreService(tempFile);
            var saved = await service.SaveAsync(null, "Local Test", SampleConnectionString, isDefault: false, CancellationToken.None);

            await service.SaveAsync(saved.Id, "Renamed", "Data Source=other;", isDefault: false, CancellationToken.None);

            var items = await service.GetAllAsync(CancellationToken.None);

            Assert.Single(items);
            Assert.Equal("Renamed", items[0].Name);
            Assert.Equal("Data Source=other;", service.Decrypt(items[0]));
        }
        finally
        {
            if (File.Exists(tempFile))
            {
                File.Delete(tempFile);
            }
        }
    }

    [Fact]
    public async Task SaveAsync_SettingDefault_ClearsDefaultOnOtherProfiles()
    {
        var tempFile = Path.Combine(Path.GetTempPath(), $"dbm-conn-{Guid.NewGuid():N}.json");

        try
        {
            var service = new ConnectionProfileStoreService(tempFile);
            var first = await service.SaveAsync(null, "First", SampleConnectionString, isDefault: true, CancellationToken.None);
            var second = await service.SaveAsync(null, "Second", SampleConnectionString, isDefault: true, CancellationToken.None);

            var items = await service.GetAllAsync(CancellationToken.None);

            Assert.Single(items, p => p.IsDefault);
            Assert.True(items.Single(p => p.Id == second.Id).IsDefault);
            Assert.False(items.Single(p => p.Id == first.Id).IsDefault);
        }
        finally
        {
            if (File.Exists(tempFile))
            {
                File.Delete(tempFile);
            }
        }
    }

    [Fact]
    public async Task DeleteAsync_RemovesProfile()
    {
        var tempFile = Path.Combine(Path.GetTempPath(), $"dbm-conn-{Guid.NewGuid():N}.json");

        try
        {
            var service = new ConnectionProfileStoreService(tempFile);
            var saved = await service.SaveAsync(null, "ToDelete", SampleConnectionString, isDefault: false, CancellationToken.None);

            await service.DeleteAsync(saved.Id, CancellationToken.None);

            var items = await service.GetAllAsync(CancellationToken.None);
            Assert.Empty(items);
        }
        finally
        {
            if (File.Exists(tempFile))
            {
                File.Delete(tempFile);
            }
        }
    }

    [Fact]
    public async Task TouchLastUsedAsync_SetsLastUsedTimestamp()
    {
        var tempFile = Path.Combine(Path.GetTempPath(), $"dbm-conn-{Guid.NewGuid():N}.json");

        try
        {
            var service = new ConnectionProfileStoreService(tempFile);
            var saved = await service.SaveAsync(null, "Local Test", SampleConnectionString, isDefault: false, CancellationToken.None);
            Assert.Null(saved.LastUsedAtUtc);

            await service.TouchLastUsedAsync(saved.Id, CancellationToken.None);

            var items = await service.GetAllAsync(CancellationToken.None);
            Assert.NotNull(items[0].LastUsedAtUtc);
        }
        finally
        {
            if (File.Exists(tempFile))
            {
                File.Delete(tempFile);
            }
        }
    }

    [Fact]
    public async Task SaveAsync_MissingNameOrConnectionString_Throws()
    {
        var tempFile = Path.Combine(Path.GetTempPath(), $"dbm-conn-{Guid.NewGuid():N}.json");
        var service = new ConnectionProfileStoreService(tempFile);

        await Assert.ThrowsAsync<ArgumentException>(
            () => service.SaveAsync(null, string.Empty, SampleConnectionString, false, CancellationToken.None));
        await Assert.ThrowsAsync<ArgumentException>(
            () => service.SaveAsync(null, "Name", string.Empty, false, CancellationToken.None));
    }
}
