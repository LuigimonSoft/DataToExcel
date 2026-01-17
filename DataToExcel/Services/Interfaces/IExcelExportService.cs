using System.Data;
using DataToExcel.Models;

namespace DataToExcel.Services.Interfaces;

public interface IExcelExportService
{
    Task<ServiceResponse<Stream>> ExportAsync(IEnumerable<IDataRecord> data,
        IReadOnlyList<ColumnDefinition> columns,
        Stream output,
        ExcelExportOptions options,
        CancellationToken ct = default);

    Task<ServiceResponse<Stream>> ExportAsync(IAsyncEnumerable<IDataRecord> data,
        IReadOnlyList<ColumnDefinition> columns,
        Stream output,
        ExcelExportOptions options,
        CancellationToken ct = default);
}
