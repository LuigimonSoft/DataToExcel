using System.Data;
using DataToExcel.Models;

namespace DataToExcel.Application.Interfaces;

public interface IExportExcel
{
    Task<IReadOnlyList<BlobUploadResult>> ExecuteAsync(IEnumerable<IDataRecord> data,
        IReadOnlyList<ColumnDefinition> columns,
        string baseFileName,
        ExcelExportOptions options,
        TimeSpan? sasTtl = null,
        CancellationToken ct = default);

    Task<IReadOnlyList<BlobUploadResult>> ExecuteAsync(IAsyncEnumerable<IDataRecord> data,
        IReadOnlyList<ColumnDefinition> columns,
        string baseFileName,
        ExcelExportOptions options,
        TimeSpan? sasTtl = null,
        CancellationToken ct = default);
}
