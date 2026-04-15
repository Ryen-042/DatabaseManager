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

            await service.ExportToCsvAsync(table, tempFile, false, CancellationToken.None);

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

            await service.ExportToExcelAsync(table, tempFile, false, CancellationToken.None);

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

    [Fact]
    public async Task ExportToCsvAsync_FormatsBinaryValuesAsHex()
    {
        var tempFile = Path.Combine(Path.GetTempPath(), $"dbm-{Guid.NewGuid():N}.csv");

        try
        {
            var service = new ExportService();
            var table = new DataTable();
            table.Columns.Add("Payload", typeof(byte[]));
            table.Rows.Add(new byte[] { 0xDE, 0xAD, 0xBE, 0xEF });

            await service.ExportToCsvAsync(table, tempFile, false, CancellationToken.None);

            var content = await File.ReadAllTextAsync(tempFile);
            Assert.Contains("Payload", content);
            Assert.Contains("0xDEADBEEF", content);
            Assert.DoesNotContain("System.Byte[]", content);
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
    public async Task ExportToCsvAsync_TruncatesLargeBinaryByDefault()
    {
        var tempFile = Path.Combine(Path.GetTempPath(), $"dbm-{Guid.NewGuid():N}.csv");

        try
        {
            var service = new ExportService();
            var table = new DataTable();
            table.Columns.Add("Payload", typeof(byte[]));

            var bytes = Enumerable.Range(0, 65).Select(i => (byte)i).ToArray();
            table.Rows.Add(bytes);

            await service.ExportToCsvAsync(table, tempFile, false, CancellationToken.None);

            var content = await File.ReadAllTextAsync(tempFile);
            Assert.Contains("... (65 bytes)", content);
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
    public async Task ExportToCsvAsync_DoesNotTruncateLargeBinaryInFullMode()
    {
        var tempFile = Path.Combine(Path.GetTempPath(), $"dbm-{Guid.NewGuid():N}.csv");

        try
        {
            var service = new ExportService();
            var table = new DataTable();
            table.Columns.Add("Payload", typeof(byte[]));

            var bytes = Enumerable.Range(0, 65).Select(i => (byte)i).ToArray();
            table.Rows.Add(bytes);

            await service.ExportToCsvAsync(table, tempFile, true, CancellationToken.None);

            var content = await File.ReadAllTextAsync(tempFile);
            Assert.Contains($"0x{Convert.ToHexString(bytes)}", content);
            Assert.DoesNotContain("... (65 bytes)", content);
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
