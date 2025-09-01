using Azure.Storage.Blobs.Models;
using Azure.Storage.Sas;
using DataToExcel.Models;
using DataToExcel.Repositories;
using DataToExcel.Wrappers.Interfaces;
using Moq;
using Xunit;

namespace DataToExcel.Test.Repositories;

public class AzureBlobStorageRepositoryTests
{
    [Fact]
    public async Task GivenMockedWrappersWhenUploadExcelAsyncThenRepositoryShouldReturnSuccess()
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

        var repo = new AzureBlobStorageRepository(containerMock.Object, TimeSpan.FromMinutes(5));

        using var stream = new MemoryStream(new byte[] {1, 2, 3});

        // When
        var response = await repo.UploadExcelAsync(stream, "file.xlsx", TimeSpan.FromMinutes(1));

        // Then
        containerMock.Verify(
            c => c.CreateIfNotExistsAsync(PublicAccessType.None, It.IsAny<CancellationToken>()),
            Times.Once);
        blobMock.Verify(
            b => b.UploadAsync(It.IsAny<Stream>(), It.IsAny<BlobUploadOptions>(), It.IsAny<CancellationToken>()),
            Times.Once);

        Assert.True(response.IsSuccess);
        Assert.NotNull(response.Data);
        Assert.Equal("test", response.Data!.Container);
        Assert.Equal("file.xlsx", response.Data!.BlobName);
        Assert.Equal(new Uri("https://example.com/blob?sas=1"), response.Data!.SasReadUri);
        Assert.Equal(new Uri("https://example.com/blob"), response.Data!.BlobUri);
    }
}
