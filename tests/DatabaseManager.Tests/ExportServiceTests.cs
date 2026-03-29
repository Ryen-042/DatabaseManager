using System.Data;
using DatabaseManager.Core.Services;

namespace DatabaseManager.Tests;

public sealed class ExportServiceTests
{
    [Fact]
    public async Task ExportToCsvAsync_WritesData()
    {
        var tempFile = Path.Combine(Path.GetTempPath(), $"dbm-{Guid.NewGuid():N}.csv");

        try
        {
            var service = new ExportService();
            var table = BuildSampleTable();

            await service.ExportToCsvAsync(table, tempFile, CancellationToken.None);

            Assert.True(File.Exists(tempFile));
            var content = await File.ReadAllTextAsync(tempFile);
            Assert.Contains("Id,Name", content);
            Assert.Contains("1,Alice", content);
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
    public async Task ExportToExcelAsync_WritesFile()
    {
        var tempFile = Path.Combine(Path.GetTempPath(), $"dbm-{Guid.NewGuid():N}.xlsx");

        try
        {
            var service = new ExportService();
            var table = BuildSampleTable();

            await service.ExportToExcelAsync(table, tempFile, CancellationToken.None);

            Assert.True(File.Exists(tempFile));
            var size = new FileInfo(tempFile).Length;
            Assert.True(size > 0);
        }
        finally
        {
            if (File.Exists(tempFile))
            {
                File.Delete(tempFile);
            }
        }
    }

    private static DataTable BuildSampleTable()
    {
        var table = new DataTable();
        table.Columns.Add("Id", typeof(int));
        table.Columns.Add("Name", typeof(string));

        table.Rows.Add(1, "Alice");
        table.Rows.Add(2, "Bob");

        return table;
    }
}
