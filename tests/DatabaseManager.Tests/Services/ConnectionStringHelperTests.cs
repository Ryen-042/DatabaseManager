using DatabaseManager.Core.Services;

namespace DatabaseManager.Tests.Services;

public sealed class ConnectionStringHelperTests
{
    [Fact]
    public void WithDatabase_NoExistingCatalog_AddsInitialCatalog()
    {
        var result = ConnectionStringHelper.WithDatabase("Server=myserver;Trusted_Connection=True;", "SalesDb");

        Assert.Contains("Initial Catalog=SalesDb", result);
    }

    [Fact]
    public void WithDatabase_ExistingCatalog_ReplacesIt()
    {
        var result = ConnectionStringHelper.WithDatabase("Server=myserver;Initial Catalog=OldDb;Trusted_Connection=True;", "NewDb");

        Assert.Contains("Initial Catalog=NewDb", result);
        Assert.DoesNotContain("OldDb", result);
    }

    [Fact]
    public void WithFallbackDatabase_NoExistingCatalog_UsesFallback()
    {
        var result = ConnectionStringHelper.WithFallbackDatabase("Server=myserver;Trusted_Connection=True;", "master");

        Assert.Contains("Initial Catalog=master", result);
    }

    [Fact]
    public void WithFallbackDatabase_ExistingCatalog_LeavesItUnchanged()
    {
        var result = ConnectionStringHelper.WithFallbackDatabase("Server=myserver;Initial Catalog=SalesDb;Trusted_Connection=True;", "master");

        Assert.Contains("Initial Catalog=SalesDb", result);
        Assert.DoesNotContain("master", result);
    }
}
