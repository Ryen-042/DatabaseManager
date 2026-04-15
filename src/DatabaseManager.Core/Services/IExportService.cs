using System.Data;

namespace DatabaseManager.Core.Services;

public interface IExportService
{
    Task ExportToCsvAsync(DataTable dataTable, string filePath, bool fullOutput, CancellationToken cancellationToken);

    Task ExportToExcelAsync(DataTable dataTable, string filePath, bool fullOutput, CancellationToken cancellationToken);
}
