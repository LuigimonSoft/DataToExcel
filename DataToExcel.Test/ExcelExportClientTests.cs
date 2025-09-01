using System.Data;
using Azure.Storage.Blobs.Models;
using Azure.Storage.Sas;
using DataToExcel.Models;
using DataToExcel.Wrappers.Interfaces;
using Moq;
using Xunit;

namespace DataToExcel.Test;

public class ExcelExportClientTests
{
    [Fact]
    public void GivenConnectionStringCtor_ShouldInitialize()
    {
        // Act
        var client = new DataToExcel.ExcelExportClient(
            connectionString: "UseDevelopmentStorage=true",
            containerName: "reports",
            defaultSasTtl: TimeSpan.FromMinutes(5));

        // Assert
        Assert.NotNull(client);
    }

    [Fact]
    public void GivenUriCtor_ShouldInitialize()
    {
        // Act
        var client = new DataToExcel.ExcelExportClient(
            containerUri: new Uri("https://example.com/container"),
            credential: null,
            defaultSasTtl: TimeSpan.FromMinutes(5));

        // Assert
        Assert.NotNull(client);
    }

    [Fact]
    public async Task GivenContainerCtor_WhenExecuteAsync_ThenUploadsBlob()
    {
        // Given sample data
        var table = new DataTable();
        table.Columns.Add("Name", typeof(string));
        table.Rows.Add("Alice");
        var reader = table.CreateDataReader();
        var records = new List<IDataRecord>();
        while (reader.Read()) records.Add(reader);

        var columns = new List<ColumnDefinition> { new("Name", "Name", ColumnDataType.String) };
        var options = new ExcelExportOptions { SheetName = "Names" };

        // Mocks for blob interactions
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

        var client = new DataToExcel.ExcelExportClient(containerMock.Object, TimeSpan.FromMinutes(5));

        // When
        var result = await client.ExecuteAsync(records, columns, "Report", options);

        // Then
        blobMock.Verify(
            b => b.UploadAsync(It.IsAny<Stream>(), It.IsAny<BlobUploadOptions>(), It.IsAny<CancellationToken>()),
            Times.Once);
        Assert.NotNull(result);
        Assert.EndsWith(".xlsx", result.BlobName);
    }
}

