using System.Data;
using System.Runtime.CompilerServices;
using DataToExcel.Application.Interfaces;
using DataToExcel.Models;
using DataToExcel.Repositories.Interfaces;
using DataToExcel.Services.Interfaces;
using DataToExcel.Utilities;

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

    public Task<IReadOnlyList<BlobUploadResult>> ExecuteAsync(IEnumerable<IDataRecord> data,
        IReadOnlyList<ColumnDefinition> columns,
        string baseFileName,
        ExcelExportOptions options,
        TimeSpan? sasTtl = null,
        CancellationToken ct = default)
        => ExecuteAsync(AsyncEnumerableHelpers.ToAsyncEnumerable(data, ct), columns, baseFileName, options, sasTtl, ct);

    public Task<IReadOnlyList<BlobUploadResult>> ExecuteAsync(IAsyncEnumerable<IDataRecord> data,
        IReadOnlyList<ColumnDefinition> columns,
        string baseFileName,
        ExcelExportOptions options,
        TimeSpan? sasTtl = null,
        CancellationToken ct = default)
        => ExecuteAsyncCore(
            options,
            () => ExecuteMultipleFileExportsAsync(data, columns, baseFileName, options, sasTtl, ct),
            () => ExecuteSingleExportAsync(
                baseFileName,
                null,
                options,
                sasTtl,
                stream => _excelService.ExportAsync(data, columns, stream, options, ct),
                appendFileIndex: false,
                fileIndex: 1,
                ct: ct));

    private static async Task<IReadOnlyList<BlobUploadResult>> ExecuteAsyncCore(
        ExcelExportOptions options,
        Func<Task<IReadOnlyList<BlobUploadResult>>> executeMultiFileAsync,
        Func<Task<BlobUploadResult>> executeSingleAsync)
    {
        if (options.SplitIntoMultipleFiles)
        {
            return await executeMultiFileAsync();
        }

        var result = await executeSingleAsync();
        return new[] { result };
    }

    private async Task<IReadOnlyList<BlobUploadResult>> ExecuteMultipleFileExportsAsync(
        IAsyncEnumerable<IDataRecord> data,
        IReadOnlyList<ColumnDefinition> columns,
        string baseFileName,
        ExcelExportOptions options,
        TimeSpan? sasTtl,
        CancellationToken ct)
    {
        await using var enumerator = data.GetAsyncEnumerator(ct);
        var bufferedEnumerator = new BufferedAsyncRecordEnumerator(enumerator);
        var exportOptions = CloneOptions(options, splitIntoMultipleSheets: false, splitIntoMultipleFiles: false);
        var baseGeneratedName = BuildBaseGeneratedName(baseFileName, options);
        return await ExecuteMultipleFileExportsAsyncCore(
            baseGeneratedName,
            sasTtl,
            () =>
            {
                var chunk = TakeNext(bufferedEnumerator, ExcelExportLimits.MaxDataRowsPerSheet, ct);
                return ExportToTempFileAsync(stream => _excelService.ExportAsync(chunk, columns, stream, exportOptions, ct), ct);
            },
            () => bufferedEnumerator.TryPeekNextAsync(),
            ct);
    }

    private async Task<IReadOnlyList<BlobUploadResult>> ExecuteMultipleFileExportsAsyncCore(
        string baseGeneratedName,
        TimeSpan? sasTtl,
        Func<Task<FileStream>> exportChunkAsync,
        Func<Task<bool>> hasMoreAsync,
        CancellationToken ct)
    {
        var exports = new List<BlobUploadResult>();
        var fileIndex = 1;

        while (true)
        {
            var exportResponse = await exportChunkAsync();
            var hasMore = await hasMoreAsync();
            var appendFileIndex = hasMore || fileIndex > 1;
            var result = await UploadExportAsync(
                exportResponse,
                baseGeneratedName,
                sasTtl,
                appendFileIndex,
                fileIndex,
                ct);
            exports.Add(result);
            fileIndex++;
            if (!hasMore)
                break;
        }

        return exports;
    }

    private async Task<BlobUploadResult> ExecuteSingleExportAsync(
        string baseFileName,
        string? baseGeneratedName,
        ExcelExportOptions options,
        TimeSpan? sasTtl,
        Func<Stream, Task<ServiceResponse<Stream>>> export,
        bool appendFileIndex,
        int fileIndex,
        CancellationToken ct)
    {
        var (dataDate, created) = ResolveDates(options);
        var fileNameBase = baseGeneratedName ?? BuildBaseFileName(baseFileName, dataDate, created);
        var fileName = appendFileIndex
            ? AppendFileIndex(fileNameBase, fileIndex)
            : fileNameBase;
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
                return await UploadExportAsync(fs, blobName, sasTtl, ct);
            }
        }
        finally
        {
            if (File.Exists(tempFile))
                File.Delete(tempFile);
        }
    }

    private static async IAsyncEnumerable<IDataRecord> TakeNext(BufferedAsyncRecordEnumerator enumerator, int maxRows,
        [EnumeratorCancellation] CancellationToken ct)
    {
        var written = 0;
        while (written < maxRows && await enumerator.TryGetNextAsync())
        {
            ct.ThrowIfCancellationRequested();
            var record = enumerator.Current ?? throw new InvalidOperationException("Expected record instance.");
            yield return record;
            written++;
        }
    }

    private static ExcelExportOptions CloneOptions(ExcelExportOptions options, bool splitIntoMultipleSheets, bool splitIntoMultipleFiles)
        => new()
        {
            SheetName = options.SheetName,
            Culture = options.Culture,
            FreezeHeader = options.FreezeHeader,
            AutoFilter = options.AutoFilter,
            DataDateUtc = options.DataDateUtc,
            SplitIntoMultipleSheets = splitIntoMultipleSheets,
            SplitIntoMultipleFiles = splitIntoMultipleFiles
        };

    private static string AppendFileIndex(string fileName, int fileIndex)
    {
        var extension = Path.GetExtension(fileName);
        var name = Path.GetFileNameWithoutExtension(fileName);
        return $"{name}_part{fileIndex:D2}{extension}";
    }

    private (DateTime dataDate, DateTime created) ResolveDates(ExcelExportOptions options)
    {
        var created = DateTime.UtcNow;
        var dataDate = options.DataDateUtc ?? created.Date;
        return (dataDate, created);
    }

    private string BuildBaseGeneratedName(string baseFileName, ExcelExportOptions options)
    {
        var (dataDate, created) = ResolveDates(options);
        return BuildBaseFileName(baseFileName, dataDate, created);
    }

    private string BuildBaseFileName(string baseFileName, DateTime dataDate, DateTime created)
    {
        var nameResponse = _namingService.ComposeExcelFileName(baseFileName, dataDate, created);
        if (!nameResponse.IsSuccess || nameResponse.Data is null)
            throw new InvalidOperationException(nameResponse.ErrorMessage ?? "File name generation failed");
        return nameResponse.Data;
    }

    private async Task<FileStream> ExportToTempFileAsync(
        Func<Stream, Task<ServiceResponse<Stream>>> export,
        CancellationToken ct)
    {
        var tempFile = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName());
        var fs = new FileStream(tempFile, FileMode.Create, FileAccess.ReadWrite, FileShare.None, 4096, FileOptions.SequentialScan);
        try
        {
            var exportResponse = await export(fs);
            if (!exportResponse.IsSuccess)
                throw new InvalidOperationException(exportResponse.ErrorMessage ?? "Excel export failed");
            fs.Position = 0;
            return fs;
        }
        catch
        {
            await fs.DisposeAsync();
            if (File.Exists(tempFile))
                File.Delete(tempFile);
            throw;
        }
    }

    private async Task<BlobUploadResult> UploadExportAsync(
        FileStream stream,
        string baseGeneratedName,
        TimeSpan? sasTtl,
        bool appendFileIndex,
        int fileIndex,
        CancellationToken ct)
    {
        var fileName = appendFileIndex
            ? AppendFileIndex(baseGeneratedName, fileIndex)
            : baseGeneratedName;
        var blobName = ComposeBlobName(_registrationOptions.BlobPrefix, fileName);
        try
        {
            return await UploadExportAsync(stream, blobName, sasTtl, ct);
        }
        finally
        {
            var path = stream.Name;
            await stream.DisposeAsync();
            if (File.Exists(path))
                File.Delete(path);
        }
    }

    private async Task<BlobUploadResult> UploadExportAsync(
        Stream stream,
        string blobName,
        TimeSpan? sasTtl,
        CancellationToken ct)
    {
        var response = await _blobRepository.UploadExcelAsync(stream, blobName, sasTtl, ct);
        if (!response.IsSuccess || response.Data is null)
            throw new InvalidOperationException(response.ErrorMessage ?? "Blob upload failed");
        return response.Data;
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
