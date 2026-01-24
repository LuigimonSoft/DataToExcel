using System.Collections;
using System.Data;
using System.Linq;
using Azure.Storage.Blobs.Models;
using Azure.Storage.Sas;
using DataToExcel.Application;
using DataToExcel.Application.Interfaces;
using DataToExcel.Models;
using DataToExcel.Repositories;
using DataToExcel.Services;
using DataToExcel.Wrappers.Interfaces;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Moq;
using Xunit;

namespace DataToExcel.Test.Application;

public class ExportExcelTests
{
    [Theory]
    [InlineData("folder/", "folder/")]
    [InlineData("folder\\subfolder", "folder/subfolder/")]
    [InlineData("///", null)]
    [InlineData("   ", null)]
    [InlineData("multi/level/folder", "multi/level/folder/")]
    public async Task GivenBlobPrefixVariantsWhenExecuteAsyncThenBlobNameIsNormalized(string prefix, string? expectedPrefix)
    {
        var (containerMock, useCase, records, columns, options) = BuildUseCase(prefix);

        var result = (await useCase.ExecuteAsync(records, columns, "Report", options)).Single();

        if (expectedPrefix is null)
        {
            containerMock.Verify(
                c => c.GetBlobClient(It.Is<string>(name => !name.Contains('/', StringComparison.Ordinal))),
                Times.Once);
        }
        else
        {
            containerMock.Verify(
                c => c.GetBlobClient(It.Is<string>(name => name.StartsWith(expectedPrefix, StringComparison.Ordinal))),
                Times.Once);
            Assert.StartsWith(expectedPrefix, result.BlobName, StringComparison.Ordinal);
        }

        Assert.EndsWith(".xlsx", result.BlobName);
    }

    [Fact]
    public async Task GivenBlobPrefixWhenExecuteAsyncThenBlobShouldUseFolder()
    {
        // Given
        var table = new DataTable();
        table.Columns.Add("Name", typeof(string));
        table.Rows.Add("Alice");
        var reader = table.CreateDataReader();
        var records = new List<IDataRecord>();
        while (reader.Read()) records.Add(reader);

        var columns = new List<ColumnDefinition> { new("Name", "Name", ColumnDataType.String) };
        var options = new ExcelExportOptions { SheetName = "Names" };

        var containerMock = new Mock<IBlobContainerClient>();
        var blobMock = new Mock<IBlobClient>();

        containerMock.Setup(c => c.Name).Returns("test");
        containerMock
            .Setup(c => c.CreateIfNotExistsAsync(PublicAccessType.None, It.IsAny<CancellationToken>()))
            .Returns(Task.CompletedTask);
        containerMock
            .Setup(c => c.GetBlobClient(It.IsAny<string>()))
            .Returns(blobMock.Object);

        blobMock.Setup(b => b.CanGenerateSasUri).Returns(true);
        blobMock.Setup(b => b.Uri).Returns(new Uri("https://example.com/blob"));
        blobMock
            .Setup(b => b.GenerateSasUri(It.IsAny<BlobSasBuilder>()))
            .Returns(new Uri("https://example.com/blob?sas=1"));
        blobMock
            .Setup(b => b.UploadAsync(It.IsAny<Stream>(), It.IsAny<BlobUploadOptions>(), It.IsAny<CancellationToken>()))
            .Returns(Task.CompletedTask);

        var repo = new AzureBlobStorageRepository(containerMock.Object, TimeSpan.FromMinutes(5));
        var registrationOptions = new ExcelExportRegistrationOptions { BlobPrefix = "exports/reports" };
        IExportExcel useCase = new ExportExcel(
            new ExcelExportService(new ExcelStyleProvider()),
            new FileNamingService(),
            repo,
            registrationOptions);

        // When
        var result = (await useCase.ExecuteAsync(records, columns, "Report", options)).Single();

        // Then
        containerMock.Verify(
            c => c.GetBlobClient(It.Is<string>(name => name.StartsWith("exports/reports/", StringComparison.Ordinal))),
            Times.Once);
        blobMock.Verify(
            b => b.UploadAsync(It.IsAny<Stream>(), It.IsAny<BlobUploadOptions>(), It.IsAny<CancellationToken>()),
            Times.Once);
        Assert.NotNull(result);
        Assert.StartsWith("exports/reports/", result.BlobName, StringComparison.Ordinal);
        Assert.EndsWith(".xlsx", result.BlobName);
    }

    [Fact]
    public async Task GivenNoBlobPrefixWhenExecuteAsyncThenBlobShouldUseContainerRoot()
    {
        // Given
        var table = new DataTable();
        table.Columns.Add("Name", typeof(string));
        table.Rows.Add("Alice");
        var reader = table.CreateDataReader();
        var records = new List<IDataRecord>();
        while (reader.Read()) records.Add(reader);

        var columns = new List<ColumnDefinition> { new("Name", "Name", ColumnDataType.String) };
        var options = new ExcelExportOptions { SheetName = "Names" };

        var containerMock = new Mock<IBlobContainerClient>();
        var blobMock = new Mock<IBlobClient>();

        containerMock.Setup(c => c.Name).Returns("test");
        containerMock
            .Setup(c => c.CreateIfNotExistsAsync(PublicAccessType.None, It.IsAny<CancellationToken>()))
            .Returns(Task.CompletedTask);
        containerMock
            .Setup(c => c.GetBlobClient(It.IsAny<string>()))
            .Returns(blobMock.Object);

        blobMock.Setup(b => b.CanGenerateSasUri).Returns(true);
        blobMock.Setup(b => b.Uri).Returns(new Uri("https://example.com/blob"));
        blobMock
            .Setup(b => b.GenerateSasUri(It.IsAny<BlobSasBuilder>()))
            .Returns(new Uri("https://example.com/blob?sas=1"));
        blobMock
            .Setup(b => b.UploadAsync(It.IsAny<Stream>(), It.IsAny<BlobUploadOptions>(), It.IsAny<CancellationToken>()))
            .Returns(Task.CompletedTask);

        var repo = new AzureBlobStorageRepository(containerMock.Object, TimeSpan.FromMinutes(5));
        var registrationOptions = new ExcelExportRegistrationOptions();
        IExportExcel useCase = new ExportExcel(
            new ExcelExportService(new ExcelStyleProvider()),
            new FileNamingService(),
            repo,
            registrationOptions);

        // When
        var result = (await useCase.ExecuteAsync(records, columns, "Report", options)).Single();

        // Then
        containerMock.Verify(
            c => c.GetBlobClient(It.Is<string>(name => !name.Contains('/', StringComparison.Ordinal))),
            Times.Once);
        blobMock.Verify(
            b => b.UploadAsync(It.IsAny<Stream>(), It.IsAny<BlobUploadOptions>(), It.IsAny<CancellationToken>()),
            Times.Once);
        Assert.NotNull(result);
        Assert.EndsWith(".xlsx", result.BlobName);
    }

    [Fact]
    public async Task GivenAsyncRecordsWhenExecuteAsyncThenBlobNameIsGenerated()
    {
        var (containerMock, useCase, syncRecords, columns, options) = BuildUseCase("async/reports");
        var records = ToAsyncEnumerable(syncRecords);

        var result = (await useCase.ExecuteAsync(records, columns, "Report", options)).Single();

        containerMock.Verify(
            c => c.GetBlobClient(It.Is<string>(name => name.StartsWith("async/reports/", StringComparison.Ordinal))),
            Times.Once);
        Assert.StartsWith("async/reports/", result.BlobName, StringComparison.Ordinal);
        Assert.EndsWith(".xlsx", result.BlobName);
    }

    [Fact]
    public async Task GivenGroupedColumnsWhenExecuteAsyncThenGroupedRowsAreWritten()
    {
        var table = BuildGroupedTable();
        var records = ToRecords(table);
        var columns = new List<ColumnDefinition>
        {
            new("Category","Category", ColumnDataType.String, Group: true),
            new("Amount","Amount", ColumnDataType.Number)
        };
        var options = new ExcelExportOptions { SheetName = "Grouped" };

        var containerMock = new Mock<IBlobContainerClient>();
        var blobMock = new Mock<IBlobClient>();
        var captured = new MemoryStream();

        containerMock.Setup(c => c.Name).Returns("test");
        containerMock
            .Setup(c => c.CreateIfNotExistsAsync(PublicAccessType.None, It.IsAny<CancellationToken>()))
            .Returns(Task.CompletedTask);
        containerMock
            .Setup(c => c.GetBlobClient(It.IsAny<string>()))
            .Returns(blobMock.Object);

        blobMock.Setup(b => b.CanGenerateSasUri).Returns(true);
        blobMock.Setup(b => b.Uri).Returns(new Uri("https://example.com/blob"));
        blobMock
            .Setup(b => b.GenerateSasUri(It.IsAny<BlobSasBuilder>()))
            .Returns(new Uri("https://example.com/blob?sas=1"));
        blobMock
            .Setup(b => b.UploadAsync(It.IsAny<Stream>(), It.IsAny<BlobUploadOptions>(), It.IsAny<CancellationToken>()))
            .Callback<Stream, BlobUploadOptions, CancellationToken>((stream, _, _) =>
            {
                stream.Position = 0;
                stream.CopyTo(captured);
            })
            .Returns(Task.CompletedTask);

        var repo = new AzureBlobStorageRepository(containerMock.Object, TimeSpan.FromMinutes(5));
        var registrationOptions = new ExcelExportRegistrationOptions();
        IExportExcel useCase = new ExportExcel(
            new ExcelExportService(new ExcelStyleProvider()),
            new FileNamingService(),
            repo,
            registrationOptions);

        var result = (await useCase.ExecuteAsync(records, columns, "Report", options)).Single();

        Assert.NotNull(result);
        captured.Position = 0;
        using var doc = SpreadsheetDocument.Open(captured, false);
        var sheet = doc.WorkbookPart!.WorksheetParts.First().Worksheet;
        var rows = sheet.GetFirstChild<SheetData>()!.Elements<Row>().ToList();

        Assert.Equal(20, rows.Skip(1).Count());
        Assert.Equal(4, rows.Skip(1).Count(r => r.OutlineLevel is null));
        Assert.Equal(16, rows.Skip(1).Count(r => r.OutlineLevel?.Value == 1));
        Assert.All(rows.Skip(1).Where(r => r.OutlineLevel?.Value == 1), r =>
        {
            Assert.Equal(string.Empty, r.Elements<Cell>().First().InnerText);
        });
    }

    [Fact]
    public async Task GivenForwardOnlyAsyncRecordsWhenExecuteAsyncThenGroupedRowsAreWritten()
    {
        var table = BuildGroupedTable();
        var records = new ForwardOnlyAsyncRecords(table);
        var columns = new List<ColumnDefinition>
        {
            new("Category","Category", ColumnDataType.String, Group: true),
            new("Amount","Amount", ColumnDataType.Number)
        };
        var options = new ExcelExportOptions { SheetName = "Grouped" };

        var containerMock = new Mock<IBlobContainerClient>();
        var blobMock = new Mock<IBlobClient>();
        var captured = new MemoryStream();

        containerMock.Setup(c => c.Name).Returns("test");
        containerMock
            .Setup(c => c.CreateIfNotExistsAsync(PublicAccessType.None, It.IsAny<CancellationToken>()))
            .Returns(Task.CompletedTask);
        containerMock
            .Setup(c => c.GetBlobClient(It.IsAny<string>()))
            .Returns(blobMock.Object);

        blobMock.Setup(b => b.CanGenerateSasUri).Returns(true);
        blobMock.Setup(b => b.Uri).Returns(new Uri("https://example.com/blob"));
        blobMock
            .Setup(b => b.GenerateSasUri(It.IsAny<BlobSasBuilder>()))
            .Returns(new Uri("https://example.com/blob?sas=1"));
        blobMock
            .Setup(b => b.UploadAsync(It.IsAny<Stream>(), It.IsAny<BlobUploadOptions>(), It.IsAny<CancellationToken>()))
            .Callback<Stream, BlobUploadOptions, CancellationToken>((stream, _, _) =>
            {
                stream.Position = 0;
                stream.CopyTo(captured);
            })
            .Returns(Task.CompletedTask);

        var repo = new AzureBlobStorageRepository(containerMock.Object, TimeSpan.FromMinutes(5));
        var registrationOptions = new ExcelExportRegistrationOptions();
        IExportExcel useCase = new ExportExcel(
            new ExcelExportService(new ExcelStyleProvider()),
            new FileNamingService(),
            repo,
            registrationOptions);

        var result = (await useCase.ExecuteAsync(records, columns, "Report", options)).Single();

        Assert.NotNull(result);
        captured.Position = 0;
        using var doc = SpreadsheetDocument.Open(captured, false);
        var sheet = doc.WorkbookPart!.WorksheetParts.First().Worksheet;
        var rows = sheet.GetFirstChild<SheetData>()!.Elements<Row>().ToList();

        Assert.Equal(20, rows.Skip(1).Count());
        Assert.Equal(4, rows.Skip(1).Count(r => r.OutlineLevel is null));
        Assert.Equal(16, rows.Skip(1).Count(r => r.OutlineLevel?.Value == 1));
        Assert.All(rows.Skip(1).Where(r => r.OutlineLevel?.Value == 1), r =>
        {
            Assert.Equal(string.Empty, r.Elements<Cell>().First().InnerText);
        });
    }

    [Fact]
    public async Task GivenSplitIntoMultipleFilesWhenRowLimitExceededThenUploadsMultipleFiles()
    {
        var records = new LargeRecordEnumerable(ExcelExportLimits.MaxDataRowsPerSheet + 1);
        var columns = new List<ColumnDefinition> { new("Value", "Value", ColumnDataType.String) };
        var options = new ExcelExportOptions
        {
            SplitIntoMultipleFiles = true,
            SplitIntoMultipleSheets = true
        };

        var blobNames = new List<string>();
        var repoMock = new Mock<DataToExcel.Repositories.Interfaces.IBlobStorageRepository>();
        repoMock
            .Setup(r => r.UploadExcelAsync(It.IsAny<Stream>(), It.IsAny<string>(), It.IsAny<TimeSpan?>(), It.IsAny<CancellationToken>()))
            .ReturnsAsync((Stream _, string blobName, TimeSpan? __, CancellationToken ___) =>
            {
                blobNames.Add(blobName);
                return new RepositoryResponse<BlobUploadResult>(
                    new BlobUploadResult("test", blobName, new Uri("https://example.com/blob"), new Uri("https://example.com/blob?sas=1"), 0));
            });

        var useCase = new ExportExcel(
            new CountingExcelExportService(),
            new FileNamingService(),
            repoMock.Object,
            new ExcelExportRegistrationOptions());

        var results = await useCase.ExecuteAsync(records, columns, "Report", options);

        Assert.Equal(2, results.Count);
        Assert.Equal(2, blobNames.Count);
        Assert.Contains(blobNames, name => name.Contains("_part01", StringComparison.Ordinal));
        Assert.Contains(blobNames, name => name.Contains("_part02", StringComparison.Ordinal));
    }

    private static (Mock<IBlobContainerClient> containerMock,
        IExportExcel useCase,
        List<IDataRecord> records,
        List<ColumnDefinition> columns,
        ExcelExportOptions options) BuildUseCase(string? blobPrefix)
    {
        var table = new DataTable();
        table.Columns.Add("Name", typeof(string));
        table.Rows.Add("Alice");
        var reader = table.CreateDataReader();
        var records = new List<IDataRecord>();
        while (reader.Read()) records.Add(reader);

        var columns = new List<ColumnDefinition> { new("Name", "Name", ColumnDataType.String) };
        var options = new ExcelExportOptions { SheetName = "Names" };

        var containerMock = new Mock<IBlobContainerClient>();
        var blobMock = new Mock<IBlobClient>();

        containerMock.Setup(c => c.Name).Returns("test");
        containerMock
            .Setup(c => c.CreateIfNotExistsAsync(PublicAccessType.None, It.IsAny<CancellationToken>()))
            .Returns(Task.CompletedTask);
        containerMock
            .Setup(c => c.GetBlobClient(It.IsAny<string>()))
            .Returns(blobMock.Object);

        blobMock.Setup(b => b.CanGenerateSasUri).Returns(true);
        blobMock.Setup(b => b.Uri).Returns(new Uri("https://example.com/blob"));
        blobMock
            .Setup(b => b.GenerateSasUri(It.IsAny<BlobSasBuilder>()))
            .Returns(new Uri("https://example.com/blob?sas=1"));
        blobMock
            .Setup(b => b.UploadAsync(It.IsAny<Stream>(), It.IsAny<BlobUploadOptions>(), It.IsAny<CancellationToken>()))
            .Returns(Task.CompletedTask);

        var repo = new AzureBlobStorageRepository(containerMock.Object, TimeSpan.FromMinutes(5));
        var registrationOptions = new ExcelExportRegistrationOptions { BlobPrefix = blobPrefix };
        var useCase = new ExportExcel(
            new ExcelExportService(new ExcelStyleProvider()),
            new FileNamingService(),
            repo,
            registrationOptions);

        return (containerMock, useCase, records, columns, options);
    }

    private static async IAsyncEnumerable<IDataRecord> ToAsyncEnumerable(IEnumerable<IDataRecord> records)
    {
        foreach (var record in records)
        {
            yield return record;
            await Task.Yield();
        }
    }

    private static DataTable BuildGroupedTable()
    {
        var table = new DataTable();
        table.Columns.Add("Category", typeof(string));
        table.Columns.Add("Amount", typeof(int));

        var groups = new[] { "A", "B", "C", "D" };
        foreach (var group in groups)
        {
            for (var i = 1; i <= 5; i++)
            {
                table.Rows.Add(group, ((Array.IndexOf(groups, group) * 5) + i) * 10);
            }
        }

        return table;
    }

    private static IEnumerable<IDataRecord> ToRecords(DataTable table)
    {
        var reader = table.CreateDataReader();
        while (reader.Read())
        {
            yield return reader;
        }
    }

    private sealed class LargeRecordEnumerable : IEnumerable<IDataRecord>
    {
        private readonly int _count;
        private readonly IDataRecord _record = new FakeDataRecord();

        public LargeRecordEnumerable(int count)
            => _count = count;

        public IEnumerator<IDataRecord> GetEnumerator()
        {
            for (var i = 0; i < _count; i++)
            {
                yield return _record;
            }
        }

        IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
    }

    private sealed class FakeDataRecord : IDataRecord
    {
        public int FieldCount => 1;
        public object this[int i] => "Value";
        public object this[string name] => "Value";

        public bool GetBoolean(int i) => false;
        public byte GetByte(int i) => 0;
        public long GetBytes(int i, long fieldOffset, byte[]? buffer, int bufferoffset, int length) => 0;
        public char GetChar(int i) => 'V';
        public long GetChars(int i, long fieldoffset, char[]? buffer, int bufferoffset, int length) => 0;
        public IDataReader GetData(int i) => throw new NotSupportedException();
        public string GetDataTypeName(int i) => "string";
        public DateTime GetDateTime(int i) => DateTime.MinValue;
        public decimal GetDecimal(int i) => 0;
        public double GetDouble(int i) => 0;
        public Type GetFieldType(int i) => typeof(string);
        public float GetFloat(int i) => 0;
        public Guid GetGuid(int i) => Guid.Empty;
        public short GetInt16(int i) => 0;
        public int GetInt32(int i) => 0;
        public long GetInt64(int i) => 0;
        public string GetName(int i) => "Value";
        public int GetOrdinal(string name) => 0;
        public string GetString(int i) => "Value";
        public object GetValue(int i) => "Value";
        public int GetValues(object[] values)
        {
            values[0] = "Value";
            return 1;
        }
        public bool IsDBNull(int i) => false;
    }

    private sealed class CountingExcelExportService : DataToExcel.Services.Interfaces.IExcelExportService
    {
        public Task<ServiceResponse<Stream>> ExportAsync(IEnumerable<IDataRecord> data,
            IReadOnlyList<ColumnDefinition> columns,
            Stream output,
            ExcelExportOptions options,
            CancellationToken ct = default)
        {
            foreach (var _ in data)
            {
                ct.ThrowIfCancellationRequested();
            }

            return Task.FromResult(new ServiceResponse<Stream>(output) { IsSuccess = true });
        }

        public async Task<ServiceResponse<Stream>> ExportAsync(IAsyncEnumerable<IDataRecord> data,
            IReadOnlyList<ColumnDefinition> columns,
            Stream output,
            ExcelExportOptions options,
            CancellationToken ct = default)
        {
            await foreach (var _ in data.WithCancellation(ct))
            {
                ct.ThrowIfCancellationRequested();
            }

            return new ServiceResponse<Stream>(output) { IsSuccess = true };
        }
    }

    private sealed class ForwardOnlyAsyncRecords : IAsyncEnumerable<IDataRecord>, IAsyncEnumerator<IDataRecord>
    {
        private readonly DataTable _table;
        private DataTableReader? _reader;
        private bool _started;

        public ForwardOnlyAsyncRecords(DataTable table)
        {
            _table = table;
        }

        public IDataRecord Current => _reader ?? throw new InvalidOperationException("Enumerator not started.");

        public IAsyncEnumerator<IDataRecord> GetAsyncEnumerator(CancellationToken cancellationToken = default)
        {
            if (_started)
            {
                throw new InvalidOperationException("This enumerator can only be iterated once.");
            }

            _started = true;
            _reader = _table.CreateDataReader();
            return this;
        }

        public ValueTask DisposeAsync()
        {
            _reader?.Dispose();
            _reader = null;
            return ValueTask.CompletedTask;
        }

        public ValueTask<bool> MoveNextAsync()
        {
            if (_reader is null)
            {
                throw new InvalidOperationException("Enumerator not started.");
            }

            return new ValueTask<bool>(_reader.Read());
        }
    }
}
