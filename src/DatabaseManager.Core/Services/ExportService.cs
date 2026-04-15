using System.Data;
using ClosedXML.Excel;
using CsvHelper;
using System.Globalization;

namespace DatabaseManager.Core.Services;

public sealed class ExportService : IExportService
{
    public async Task ExportToCsvAsync(DataTable dataTable, string filePath, bool fullOutput, CancellationToken cancellationToken)
    {
        await using var stream = new FileStream(filePath, FileMode.Create, FileAccess.Write, FileShare.None);
        await using var writer = new StreamWriter(stream);
        await using var csv = new CsvWriter(writer, CultureInfo.InvariantCulture);

        foreach (DataColumn column in dataTable.Columns)
        {
            csv.WriteField(column.ColumnName);
        }

        await csv.NextRecordAsync();

        foreach (DataRow row in dataTable.Rows)
        {
            cancellationToken.ThrowIfCancellationRequested();

            foreach (DataColumn column in dataTable.Columns)
            {
                csv.WriteField(DisplayValueFormatter.FormatForDisplay(row[column], fullOutput));
            }

            await csv.NextRecordAsync();
        }

        await writer.FlushAsync(cancellationToken);
    }

    public Task ExportToExcelAsync(DataTable dataTable, string filePath, bool fullOutput, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();

        using var workbook = new XLWorkbook();
        var worksheet = workbook.Worksheets.Add("Results");

        for (var i = 0; i < dataTable.Columns.Count; i++)
        {
            worksheet.Cell(1, i + 1).Value = dataTable.Columns[i].ColumnName;
            worksheet.Cell(1, i + 1).Style.Font.Bold = true;
        }

        for (var rowIndex = 0; rowIndex < dataTable.Rows.Count; rowIndex++)
        {
            cancellationToken.ThrowIfCancellationRequested();

            for (var columnIndex = 0; columnIndex < dataTable.Columns.Count; columnIndex++)
            {
                worksheet.Cell(rowIndex + 2, columnIndex + 1).Value = DisplayValueFormatter.FormatForDisplay(dataTable.Rows[rowIndex][columnIndex], fullOutput);
            }
        }

        worksheet.Columns().AdjustToContents();
        workbook.SaveAs(filePath);

        return Task.CompletedTask;
    }
}
