using System.Data;
using System.Runtime.CompilerServices;
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

    public async Task<IReadOnlyList<BlobUploadResult>> ExecuteAsync(IEnumerable<IDataRecord> data,
        IReadOnlyList<ColumnDefinition> columns,
        string baseFileName,
        ExcelExportOptions options,
        TimeSpan? sasTtl = null,
        CancellationToken ct = default)
    {
        if (options.SplitIntoMultipleFiles)
        {
            return await ExecuteMultipleFileExportsAsync(data, columns, baseFileName, options, sasTtl, ct);
        }

        var result = await ExecuteSingleExportAsync(
            baseFileName,
            null,
            options,
            sasTtl,
            stream => _excelService.ExportAsync(data, columns, stream, options, ct),
            ct,
            appendFileIndex: false,
            fileIndex: 1);
        return new[] { result };
    }

    public async Task<IReadOnlyList<BlobUploadResult>> ExecuteAsync(IAsyncEnumerable<IDataRecord> data,
        IReadOnlyList<ColumnDefinition> columns,
        string baseFileName,
        ExcelExportOptions options,
        TimeSpan? sasTtl = null,
        CancellationToken ct = default)
    {
        if (options.SplitIntoMultipleFiles)
        {
            return await ExecuteMultipleFileExportsAsync(data, columns, baseFileName, options, sasTtl, ct);
        }

        var result = await ExecuteSingleExportAsync(
            baseFileName,
            null,
            options,
            sasTtl,
            stream => _excelService.ExportAsync(data, columns, stream, options, ct),
            ct,
            appendFileIndex: false,
            fileIndex: 1);
        return new[] { result };
    }

    private async Task<IReadOnlyList<BlobUploadResult>> ExecuteMultipleFileExportsAsync(
        IEnumerable<IDataRecord> data,
        IReadOnlyList<ColumnDefinition> columns,
        string baseFileName,
        ExcelExportOptions options,
        TimeSpan? sasTtl,
        CancellationToken ct)
    {
        using var enumerator = data.GetEnumerator();
        var bufferedEnumerator = new BufferedRecordEnumerator(enumerator);
        var exports = new List<BlobUploadResult>();
        var fileIndex = 1;
        var (dataDate, created) = ResolveDates(options);
        var baseGeneratedName = BuildBaseFileName(baseFileName, dataDate, created);

        while (true)
        {
            var chunk = TakeNext(bufferedEnumerator, ExcelExportLimits.MaxDataRowsPerSheet, ct);
            var exportOptions = CloneOptions(options, splitIntoMultipleSheets: false, splitIntoMultipleFiles: false);
            var exportResponse = await ExportToTempFileAsync(
                stream => _excelService.ExportAsync(chunk, columns, stream, exportOptions, ct),
                ct);
            var hasMore = bufferedEnumerator.TryPeekNext(out _);
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
        var exports = new List<BlobUploadResult>();
        var fileIndex = 1;
        var (dataDate, created) = ResolveDates(options);
        var baseGeneratedName = BuildBaseFileName(baseFileName, dataDate, created);

        while (true)
        {
            var chunk = TakeNext(bufferedEnumerator, ExcelExportLimits.MaxDataRowsPerSheet, ct);
            var exportOptions = CloneOptions(options, splitIntoMultipleSheets: false, splitIntoMultipleFiles: false);
            var exportResponse = await ExportToTempFileAsync(
                stream => _excelService.ExportAsync(chunk, columns, stream, exportOptions, ct),
                ct);
            var hasMore = await bufferedEnumerator.TryPeekNextAsync();
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
        CancellationToken ct,
        bool appendFileIndex,
        int fileIndex)
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

    private static IEnumerable<IDataRecord> TakeNext(BufferedRecordEnumerator enumerator, int maxRows, CancellationToken ct)
    {
        var written = 0;
        while (written < maxRows && enumerator.TryGetNext(out var record))
        {
            ct.ThrowIfCancellationRequested();
            yield return record;
            written++;
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

    private sealed class BufferedRecordEnumerator
    {
        private readonly IEnumerator<IDataRecord> _inner;
        private bool _hasBuffered;
        private IDataRecord? _buffered;

        public BufferedRecordEnumerator(IEnumerator<IDataRecord> inner)
            => _inner = inner;

        public bool TryGetNext(out IDataRecord record)
        {
            if (_hasBuffered)
            {
                record = _buffered ?? throw new InvalidOperationException("Buffered record expected.");
                _buffered = null;
                _hasBuffered = false;
                return true;
            }

            if (_inner.MoveNext())
            {
                record = _inner.Current;
                return true;
            }

            record = null!;
            return false;
        }

        public bool TryPeekNext(out IDataRecord? record)
        {
            if (_hasBuffered)
            {
                record = _buffered;
                return true;
            }

            if (_inner.MoveNext())
            {
                _buffered = _inner.Current;
                _hasBuffered = true;
                record = _buffered;
                return true;
            }

            record = null;
            return false;
        }
    }

    private sealed class BufferedAsyncRecordEnumerator
    {
        private readonly IAsyncEnumerator<IDataRecord> _inner;
        private bool _hasBuffered;
        public IDataRecord? Current { get; private set; }

        public BufferedAsyncRecordEnumerator(IAsyncEnumerator<IDataRecord> inner)
            => _inner = inner;

        public async Task<bool> TryGetNextAsync()
        {
            if (_hasBuffered)
            {
                _hasBuffered = false;
                return true;
            }

            if (await _inner.MoveNextAsync())
            {
                Current = _inner.Current;
                return true;
            }

            Current = null;
            return false;
        }

        public async Task<bool> TryPeekNextAsync()
        {
            if (_hasBuffered)
                return true;

            if (await _inner.MoveNextAsync())
            {
                Current = _inner.Current;
                _hasBuffered = true;
                return true;
            }

            Current = null;
            return false;
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
