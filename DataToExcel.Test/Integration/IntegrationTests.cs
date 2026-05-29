using System.Data;
using System.Linq;
using Azure.Storage.Blobs.Models;
using Azure.Storage.Sas;
using DataToExcel.Application.Interfaces;
using DataToExcel.Hosting;
using DataToExcel.Models;
using DataToExcel.Repositories;
using DataToExcel.Repositories.Interfaces;
using DataToExcel.Services;
using DataToExcel.Wrappers.Interfaces;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.Extensions.DependencyInjection;
using Moq;
using Xunit;

namespace DataToExcel.Test.Integration;

public class IntegrationTests
{
    private const int LargeExportRowCount = 800_000;
    private const long LargeExportMaxManagedMemoryGrowthBytes = 1536L * 1024 * 1024;
    private const long ReleasedMemoryMaxRetainedGrowthBytes = 16L * 1024 * 1024;

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
        var result = (await useCase.ExecuteAsync(records, columns, "Report", new ExcelExportOptions())).Single();

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

        var result = (await useCase.ExecuteAsync(records, columns, "Report", new ExcelExportOptions())).Single();

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

        var result = (await useCase.ExecuteAsync(records, columns, "Report", new ExcelExportOptions())).Single();

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
    public async Task GivenLargeStreamingExportWhenFileIsGeneratedThenManagedMemoryGrowthStaysBounded()
    {
        var service = new ExcelExportService(new ExcelStyleProvider());
        var columns = BuildLargeExportColumns();
        var options = new ExcelExportOptions { SheetName = "LargeExport" };
        var tempFile = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName());

        await ForceFullCollectionAsync();
        var baseline = GC.GetTotalMemory(forceFullCollection: true);

        try
        {
            await using var stream = new FileStream(tempFile, FileMode.Create, FileAccess.ReadWrite, FileShare.None);
            var result = await MeasureManagedMemoryDuringAsync(async () =>
                await service.ExportAsync(new LargeStreamingRecords(LargeExportRowCount), columns, stream, options));

            Assert.True(result.Response.IsSuccess, result.Response.ErrorMessage);
            Assert.True(stream.Length > 0);
            Assert.True(result.PeakBytes - baseline < LargeExportMaxManagedMemoryGrowthBytes,
                $"Managed memory growth was {(result.PeakBytes - baseline) / 1024 / 1024} MB.");
        }
        finally
        {
            if (File.Exists(tempFile))
                File.Delete(tempFile);
        }
    }

    [Fact]
    public async Task GivenLargeStreamingExportWhenReferencesAreReleasedThenManagedMemoryIsReclaimed()
    {
        var service = new ExcelExportService(new ExcelStyleProvider());
        var columns = BuildLargeExportColumns();
        var options = new ExcelExportOptions { SheetName = "LargeExport" };
        var tempFile = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName());

        await ForceFullCollectionAsync();
        var baseline = GC.GetTotalMemory(forceFullCollection: true);

        try
        {
            await using (var stream = new FileStream(tempFile, FileMode.Create, FileAccess.ReadWrite, FileShare.None))
            {
                var response = await service.ExportAsync(new LargeStreamingRecords(LargeExportRowCount), columns, stream, options);
                Assert.True(response.IsSuccess, response.ErrorMessage);
                Assert.True(stream.Length > 0);
            }

            service = null!;
            columns = null!;
            options = null!;
            await ForceFullCollectionAsync();

            var retainedGrowth = GC.GetTotalMemory(forceFullCollection: true) - baseline;
            Assert.True(retainedGrowth < ReleasedMemoryMaxRetainedGrowthBytes,
                $"Retained managed memory growth was {retainedGrowth / 1024 / 1024} MB.");
        }
        finally
        {
            if (File.Exists(tempFile))
                File.Delete(tempFile);
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

    private static IReadOnlyList<ColumnDefinition> BuildLargeExportColumns()
        => new List<ColumnDefinition>
        {
            new("Id", "Id", ColumnDataType.Number),
            new("Name", "Name", ColumnDataType.String),
            new("Amount", "Amount", ColumnDataType.Number),
            new("CreatedOn", "CreatedOn", ColumnDataType.DateTime),
            new("IsActive", "IsActive", ColumnDataType.Boolean),
            new("Region", "Region", ColumnDataType.String),
            new("Status", "Status", ColumnDataType.String),
            new("Category", "Category", ColumnDataType.String),
            new("Subcategory", "Subcategory", ColumnDataType.String),
            new("Quantity", "Quantity", ColumnDataType.Number),
            new("UnitPrice", "UnitPrice", ColumnDataType.Number),
            new("Discount", "Discount", ColumnDataType.Number),
            new("Tax", "Tax", ColumnDataType.Number),
            new("Total", "Total", ColumnDataType.Number),
            new("CreatedBy", "CreatedBy", ColumnDataType.String),
            new("UpdatedOn", "UpdatedOn", ColumnDataType.DateTime),
            new("UpdatedBy", "UpdatedBy", ColumnDataType.String),
            new("ExternalId", "ExternalId", ColumnDataType.String),
            new("Currency", "Currency", ColumnDataType.String),
            new("Country", "Country", ColumnDataType.String),
            new("Channel", "Channel", ColumnDataType.String),
            new("Priority", "Priority", ColumnDataType.Number),
            new("Score", "Score", ColumnDataType.Number),
            new("Reviewed", "Reviewed", ColumnDataType.Boolean),
            new("Notes", "Notes", ColumnDataType.String)
        };

    private static async Task<(ServiceResponse<Stream> Response, long PeakBytes)> MeasureManagedMemoryDuringAsync(
        Func<Task<ServiceResponse<Stream>>> exportAsync)
    {
        var peakBytes = GC.GetTotalMemory(forceFullCollection: false);
        using var cts = new CancellationTokenSource();
        var sampler = Task.Run(async () =>
        {
            while (!cts.Token.IsCancellationRequested)
            {
                peakBytes = Math.Max(peakBytes, GC.GetTotalMemory(forceFullCollection: false));
                await Task.Delay(25, cts.Token).ConfigureAwait(false);
            }
        });

        try
        {
            var response = await exportAsync();
            peakBytes = Math.Max(peakBytes, GC.GetTotalMemory(forceFullCollection: false));
            return (response, peakBytes);
        }
        finally
        {
            await cts.CancelAsync();
            try
            {
                await sampler;
            }
            catch (OperationCanceledException)
            {
            }
        }
    }

    private static async Task ForceFullCollectionAsync()
    {
        GC.Collect(GC.MaxGeneration, GCCollectionMode.Forced, blocking: true, compacting: true);
        GC.WaitForPendingFinalizers();
        GC.Collect(GC.MaxGeneration, GCCollectionMode.Forced, blocking: true, compacting: true);
        await Task.Yield();
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

    private sealed class LargeStreamingRecords : IAsyncEnumerable<IDataRecord>, IAsyncEnumerator<IDataRecord>
    {
        private readonly int _count;
        private readonly LargeStreamingRecord _record = new();
        private int _current;

        public LargeStreamingRecords(int count)
            => _count = count;

        public IDataRecord Current => _record;

        public IAsyncEnumerator<IDataRecord> GetAsyncEnumerator(CancellationToken cancellationToken = default)
            => this;

        public ValueTask DisposeAsync()
            => ValueTask.CompletedTask;

        public ValueTask<bool> MoveNextAsync()
        {
            if (_current >= _count)
                return new ValueTask<bool>(false);

            _current++;
            _record.SetRow(_current);
            return new ValueTask<bool>(true);
        }
    }

    private sealed class LargeStreamingRecord : IDataRecord
    {
        private static readonly string[] Names =
        [
            "Id", "Name", "Amount", "CreatedOn", "IsActive",
            "Region", "Status", "Category", "Subcategory", "Quantity",
            "UnitPrice", "Discount", "Tax", "Total", "CreatedBy",
            "UpdatedOn", "UpdatedBy", "ExternalId", "Currency", "Country",
            "Channel", "Priority", "Score", "Reviewed", "Notes"
        ];

        private static readonly string[] CustomerNames = Enumerable.Range(0, 1_024)
            .Select(i => $"Customer {i:D4}")
            .ToArray();

        private static readonly string[] ExternalIds = Enumerable.Range(0, 1_024)
            .Select(i => $"EXT-{i:D8}")
            .ToArray();

        private static readonly string[] Notes = Enumerable.Range(0, 1_024)
            .Select(i => $"Generated note {i:D4}")
            .ToArray();

        private static readonly string[] Regions = Enumerable.Range(0, 8)
            .Select(i => $"Region {i}")
            .ToArray();

        private static readonly string[] Categories = Enumerable.Range(0, 12)
            .Select(i => $"Category {i}")
            .ToArray();

        private static readonly string[] Subcategories = Enumerable.Range(0, 24)
            .Select(i => $"Subcategory {i}")
            .ToArray();

        private static readonly string[] CreatedByUsers = Enumerable.Range(0, 50)
            .Select(i => $"User {i}")
            .ToArray();

        private static readonly string[] UpdatedByUsers = Enumerable.Range(0, 25)
            .Select(i => $"Updater {i}")
            .ToArray();

        private int _row;

        public int FieldCount => Names.Length;
        public object this[int i] => GetValue(i);
        public object this[string name] => GetValue(GetOrdinal(name));

        public void SetRow(int row)
            => _row = row;

        public bool GetBoolean(int i) => _row % 2 == 0;
        public byte GetByte(int i) => (byte)(_row % byte.MaxValue);
        public long GetBytes(int i, long fieldOffset, byte[]? buffer, int bufferoffset, int length) => 0;
        public char GetChar(int i) => 'A';
        public long GetChars(int i, long fieldoffset, char[]? buffer, int bufferoffset, int length) => 0;
        public IDataReader GetData(int i) => throw new NotSupportedException();
        public string GetDataTypeName(int i) => GetFieldType(i).Name;
        public DateTime GetDateTime(int i) => new(2026, 1, 1);
        public decimal GetDecimal(int i) => _row / 10m;
        public double GetDouble(int i) => _row / 10d;
        public Type GetFieldType(int i)
            => i switch
            {
                0 => typeof(int),
                2 => typeof(decimal),
                3 => typeof(DateTime),
                4 => typeof(bool),
                9 => typeof(int),
                10 => typeof(decimal),
                11 => typeof(decimal),
                12 => typeof(decimal),
                13 => typeof(decimal),
                15 => typeof(DateTime),
                21 => typeof(int),
                22 => typeof(decimal),
                23 => typeof(bool),
                _ => typeof(string)
            };

        public float GetFloat(int i) => _row / 10f;
        public Guid GetGuid(int i) => Guid.Empty;
        public short GetInt16(int i) => (short)_row;
        public int GetInt32(int i) => _row;
        public long GetInt64(int i) => _row;
        public string GetName(int i) => Names[i];
        public int GetOrdinal(string name) => Array.IndexOf(Names, name);
        public string GetString(int i) => CustomerNames[_row % CustomerNames.Length];

        public object GetValue(int i)
            => i switch
            {
                0 => _row,
                1 => GetString(i),
                2 => GetDecimal(i),
                3 => GetDateTime(i),
                4 => GetBoolean(i),
                5 => Regions[_row % Regions.Length],
                6 => _row % 3 == 0 ? "Closed" : "Open",
                7 => Categories[_row % Categories.Length],
                8 => Subcategories[_row % Subcategories.Length],
                9 => _row % 100,
                10 => _row / 100m,
                11 => (_row % 15) / 100m,
                12 => (_row % 20) / 100m,
                13 => _row * 1.13m,
                14 => CreatedByUsers[_row % CreatedByUsers.Length],
                15 => GetDateTime(i).AddDays(_row % 30),
                16 => UpdatedByUsers[_row % UpdatedByUsers.Length],
                17 => ExternalIds[_row % ExternalIds.Length],
                18 => _row % 2 == 0 ? "USD" : "CAD",
                19 => _row % 2 == 0 ? "US" : "CA",
                20 => _row % 3 == 0 ? "Online" : "Retail",
                21 => _row % 5,
                22 => (_row % 1000) / 10m,
                23 => _row % 4 == 0,
                24 => Notes[_row % Notes.Length],
                _ => throw new IndexOutOfRangeException()
            };

        public int GetValues(object[] values)
        {
            var count = Math.Min(values.Length, FieldCount);
            for (var i = 0; i < count; i++)
            {
                values[i] = GetValue(i);
            }

            return count;
        }

        public bool IsDBNull(int i) => false;
    }
}
