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
    [Fact]
    public async Task GivenValidRecordsWhenExecuteAsyncThenBlobShouldBeUploaded()
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
        var useCase = new ExportExcel(
            new ExcelExportService(new ExcelStyleProvider()),
            new FileNamingService(),
            repo);

        // When
        var result = await useCase.ExecuteAsync(records, columns, "Report", options);

        // Then
        blobMock.Verify(
            b => b.UploadAsync(It.IsAny<Stream>(), It.IsAny<BlobUploadOptions>(), It.IsAny<CancellationToken>()),
            Times.Once);
        Assert.NotNull(result);
        Assert.EndsWith(".xlsx", result.BlobName);
    }
}
