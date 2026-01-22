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

        var result = await useCase.ExecuteAsync(records, columns, "Report", options);

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
        var result = await useCase.ExecuteAsync(records, columns, "Report", options);

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
        var result = await useCase.ExecuteAsync(records, columns, "Report", options);

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

        var result = await useCase.ExecuteAsync(records, columns, "Report", options);

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

        var result = await useCase.ExecuteAsync(records, columns, "Report", options);

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

        var result = await useCase.ExecuteAsync(records, columns, "Report", options);

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
