using System.Data;
using DataToExcel.Models;

namespace DataToExcel.Application.Interfaces;

public interface IExportExcel
{
    Task<BlobUploadResult> ExecuteAsync(IEnumerable<IDataRecord> data,
        IReadOnlyList<ColumnDefinition> columns,
        string baseFileName,
        ExcelExportOptions options,
        TimeSpan? sasTtl = null,
        CancellationToken ct = default);
}
