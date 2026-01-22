using System.Data;
using System.Linq;
using Azure.Storage.Blobs.Models;
using Azure.Storage.Sas;
using DataToExcel.Application.Interfaces;
using DataToExcel.Hosting;
using DataToExcel.Models;
using DataToExcel.Repositories;
using DataToExcel.Repositories.Interfaces;
using DataToExcel.Wrappers.Interfaces;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.Extensions.DependencyInjection;
using Moq;
using Xunit;

namespace DataToExcel.Test.Integration;

public class IntegrationTests
{
    [Fact]
    public async Task GivenMockedBlobStorageWhenUseCaseExecutesViaDIThenBlobShouldBeUploaded()
    {
        // Given
        var services = new ServiceCollection();
        services.AddExcelExport(o =>
        {
            o.ConnectionString = "UseDevelopmentStorage=true";
            o.ContainerName = "test";
        });

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

        services.AddSingleton<IBlobStorageRepository>(sp =>
            new AzureBlobStorageRepository(containerMock.Object, TimeSpan.FromMinutes(5)));

        var provider = services.BuildServiceProvider();
        var useCase = provider.GetRequiredService<IExportExcel>();

        var table = new DataTable();
        table.Columns.Add("Name", typeof(string));
        table.Rows.Add("Alice");
        var reader = table.CreateDataReader();
        var records = new List<IDataRecord>();
        while (reader.Read()) records.Add(reader);

        var columns = new List<ColumnDefinition> { new("Name", "Name", ColumnDataType.String) };

        // When
        var result = await useCase.ExecuteAsync(records, columns, "Report", new ExcelExportOptions());

        // Then
        blobMock.Verify(
            b => b.UploadAsync(It.IsAny<Stream>(), It.IsAny<BlobUploadOptions>(), It.IsAny<CancellationToken>()),
            Times.Once);
        Assert.Equal("test", result.Container);
    }

    [Fact]
    public async Task GivenGroupedColumnsWhenUseCaseExecutesViaDIThenGroupedRowsAreWritten()
    {
        var services = new ServiceCollection();
        services.AddExcelExport(o =>
        {
            o.ConnectionString = "UseDevelopmentStorage=true";
            o.ContainerName = "test";
        });

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

        services.AddSingleton<IBlobStorageRepository>(sp =>
            new AzureBlobStorageRepository(containerMock.Object, TimeSpan.FromMinutes(5)));

        var provider = services.BuildServiceProvider();
        var useCase = provider.GetRequiredService<IExportExcel>();

        var table = BuildGroupedTable();
        var records = ToRecords(table);
        var columns = new List<ColumnDefinition>
        {
            new("Category", "Category", ColumnDataType.String, Group: true),
            new("Amount", "Amount", ColumnDataType.Number)
        };

        var result = await useCase.ExecuteAsync(records, columns, "Report", new ExcelExportOptions());

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
    public async Task GivenForwardOnlyAsyncRecordsWhenUseCaseExecutesViaDIThenGroupedRowsAreWritten()
    {
        var services = new ServiceCollection();
        services.AddExcelExport(o =>
        {
            o.ConnectionString = "UseDevelopmentStorage=true";
            o.ContainerName = "test";
        });

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

        services.AddSingleton<IBlobStorageRepository>(sp =>
            new AzureBlobStorageRepository(containerMock.Object, TimeSpan.FromMinutes(5)));

        var provider = services.BuildServiceProvider();
        var useCase = provider.GetRequiredService<IExportExcel>();

        var table = BuildGroupedTable();
        var records = new ForwardOnlyAsyncRecords(table);
        var columns = new List<ColumnDefinition>
        {
            new("Category", "Category", ColumnDataType.String, Group: true),
            new("Amount", "Amount", ColumnDataType.Number)
        };

        var result = await useCase.ExecuteAsync(records, columns, "Report", new ExcelExportOptions());

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
