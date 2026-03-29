using System.Data;

namespace DatabaseManager.Core.Services;

public interface IExportService
{
    Task ExportToCsvAsync(DataTable dataTable, string filePath, CancellationToken cancellationToken);

    Task ExportToExcelAsync(DataTable dataTable, string filePath, CancellationToken cancellationToken);
}
