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
    private readonly ExcelExportRegistrationOptions _registrationOptions;

    public ExportExcel(IExcelExportService excelService,
        IFileNamingService namingService,
        IBlobStorageRepository blobRepository,
        ExcelExportRegistrationOptions registrationOptions)
    {
        _excelService = excelService;
        _namingService = namingService;
        _blobRepository = blobRepository;
        _registrationOptions = registrationOptions;
    }

    public async Task<BlobUploadResult> ExecuteAsync(IEnumerable<IDataRecord> data,
        IReadOnlyList<ColumnDefinition> columns,
        string baseFileName,
        ExcelExportOptions options,
        TimeSpan? sasTtl = null,
        CancellationToken ct = default)
        => await ExecuteWithExportAsync(
            baseFileName,
            options,
            sasTtl,
            ct,
            stream => _excelService.ExportAsync(data, columns, stream, options, ct));

    public async Task<BlobUploadResult> ExecuteAsync(IAsyncEnumerable<IDataRecord> data,
        IReadOnlyList<ColumnDefinition> columns,
        string baseFileName,
        ExcelExportOptions options,
        TimeSpan? sasTtl = null,
        CancellationToken ct = default)
        => await ExecuteWithExportAsync(
            baseFileName,
            options,
            sasTtl,
            ct,
            stream => _excelService.ExportAsync(data, columns, stream, options, ct));

    private async Task<BlobUploadResult> ExecuteWithExportAsync(
        string baseFileName,
        ExcelExportOptions options,
        TimeSpan? sasTtl,
        CancellationToken ct,
        Func<Stream, Task<ServiceResponse<Stream>>> export)
    {
        var created = DateTime.UtcNow;
        var dataDate = options.DataDateUtc ?? created.Date;
        var nameResponse = _namingService.ComposeExcelFileName(baseFileName, dataDate, created);
        if (!nameResponse.IsSuccess || nameResponse.Data is null)
            throw new InvalidOperationException(nameResponse.ErrorMessage ?? "File name generation failed");
        var fileName = nameResponse.Data;
        var blobName = ComposeBlobName(_registrationOptions.BlobPrefix, fileName);

        var tempFile = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName());
        try
        {
            await using (var fs = new FileStream(tempFile, FileMode.Create, FileAccess.ReadWrite, FileShare.None, 4096, FileOptions.SequentialScan))
            {
                var exportResponse = await export(fs);
                if (!exportResponse.IsSuccess)
                    throw new InvalidOperationException(exportResponse.ErrorMessage ?? "Excel export failed");
                fs.Position = 0;
                var response = await _blobRepository.UploadExcelAsync(fs, blobName, sasTtl, ct);
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

    private static string ComposeBlobName(string? prefix, string fileName)
    {
        if (string.IsNullOrWhiteSpace(prefix))
            return fileName;

        var cleaned = prefix.Replace('\\', '/').Trim('/');
        if (string.IsNullOrWhiteSpace(cleaned))
            return fileName;

        return $"{cleaned}/{fileName}";
    }
}
