using System.Data;
using DataToExcel.Application.Interfaces;
using DataToExcel.Models;
using DataToExcel.Repositories.Interfaces;
using DataToExcel.Services.Interfaces;

namespace DataToExcel.Application;

public class ExportExcel : IExportExcel
{
    private readonly IExcelExportService _excelService;
    private readonly IFileNamingService _namingService;
    private readonly IBlobStorageRepository _blobRepository;

    public ExportExcel(IExcelExportService excelService,
        IFileNamingService namingService,
        IBlobStorageRepository blobRepository)
    {
        _excelService = excelService;
        _namingService = namingService;
        _blobRepository = blobRepository;
    }

    public async Task<BlobUploadResult> ExecuteAsync(IEnumerable<IDataRecord> data,
        IReadOnlyList<ColumnDefinition> columns,
        string baseFileName,
        ExcelExportOptions options,
        TimeSpan? sasTtl = null,
        CancellationToken ct = default)
    {
        var created = DateTime.UtcNow;
        var dataDate = options.DataDateUtc ?? created.Date;
        var nameResponse = _namingService.ComposeExcelFileName(baseFileName, dataDate, created);
        if (!nameResponse.IsSuccess || nameResponse.Data is null)
            throw new InvalidOperationException(nameResponse.ErrorMessage ?? "File name generation failed");
        var fileName = nameResponse.Data;

        var tempFile = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName());
        try
        {
            await using (var fs = new FileStream(tempFile, FileMode.Create, FileAccess.ReadWrite, FileShare.None, 4096, FileOptions.SequentialScan))
            {
                var exportResponse = await _excelService.ExportAsync(data, columns, fs, options, ct);
                if (!exportResponse.IsSuccess)
                    throw new InvalidOperationException(exportResponse.ErrorMessage ?? "Excel export failed");
                fs.Position = 0;
                var response = await _blobRepository.UploadExcelAsync(fs, fileName, sasTtl, ct);
                if (!response.IsSuccess || response.Data is null)
                    throw new InvalidOperationException(response.ErrorMessage ?? "Blob upload failed");
                return response.Data;
            }
        }
        finally
        {
            if (File.Exists(tempFile))
                File.Delete(tempFile);
        }
    }
}
