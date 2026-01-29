using System.Data;
using System.Linq;
using Azure.Storage.Blobs.Models;
using Azure.Storage.Sas;
using DataToExcel;
using DataToExcel.Models;
using DataToExcel.Wrappers.Interfaces;
using Moq;
using Xunit;

namespace DataToExcel.Test.Integration;

public class ExcelExportClientTests
{
    [Fact]
    public async Task GivenMockedBlobStorageWhenExecuteAsyncThenBlobShouldBeUploaded()
    {
        // Given
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

        var client = new ExcelExportClient(containerMock.Object);

        var table = new DataTable();
        table.Columns.Add("Name", typeof(string));
        table.Rows.Add("Alice");
        var reader = table.CreateDataReader();
        var records = new List<IDataRecord>();
        while (reader.Read()) records.Add(reader);

        var columns = new List<ColumnDefinition> { new("Name", "Name", ColumnDataType.String) };

        // When
        var result = (await client.ExecuteAsync(records, columns, "Report", new ExcelExportOptions())).Single();

        // Then
        blobMock.Verify(
            b => b.UploadAsync(It.IsAny<Stream>(), It.IsAny<BlobUploadOptions>(), It.IsAny<CancellationToken>()),
            Times.Once);
        Assert.Equal("test", result.Container);
    }
}
