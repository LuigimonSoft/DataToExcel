using System.Data;
using DataToExcel.Application;
using Azure.Storage.Blobs.Models;
using Azure.Storage.Sas;
using DataToExcel.Models;
using DataToExcel.Repositories;
using DataToExcel.Services;
using DataToExcel.Wrappers.Interfaces;
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
        var useCase = new ExportExcel(
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
        var useCase = new ExportExcel(
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

    private static (Mock<IBlobContainerClient> containerMock,
        ExportExcel useCase,
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
}
