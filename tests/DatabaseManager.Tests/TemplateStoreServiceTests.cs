using DatabaseManager.Core.Models;
using DatabaseManager.Core.Services;

namespace DatabaseManager.Tests;

public sealed class TemplateStoreServiceTests
{
    [Fact]
    public async Task SaveAsync_ThenGetAllAsync_ReturnsTemplate()
    {
        var tempFile = Path.Combine(Path.GetTempPath(), $"dbm-{Guid.NewGuid():N}.json");

        try
        {
            var service = new TemplateStoreService(tempFile);

            await service.SaveAsync(new QueryTemplate
            {
                Name = "AllCustomers",
                Sql = "SELECT * FROM Customers"
            }, CancellationToken.None);

            var items = await service.GetAllAsync(CancellationToken.None);

            Assert.Single(items);
            Assert.Equal("AllCustomers", items[0].Name);
            Assert.Equal("SELECT * FROM Customers", items[0].Sql);
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
    public async Task DeleteAsync_RemovesTemplate()
    {
        var tempFile = Path.Combine(Path.GetTempPath(), $"dbm-{Guid.NewGuid():N}.json");

        try
        {
            var service = new TemplateStoreService(tempFile);

            await service.SaveAsync(new QueryTemplate
            {
                Name = "ToDelete",
                Sql = "SELECT 1"
            }, CancellationToken.None);

            await service.DeleteAsync("ToDelete", CancellationToken.None);

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
}
