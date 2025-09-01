using DataToExcel.Models;
using DataToExcel.Repositories;
using DataToExcel.Wrappers.Interfaces;
using Azure.Storage.Blobs.Models;
using Azure.Storage.Sas;
using Moq;
using Xunit;

namespace DataToExcel.Test.Repositories;

public class AzureBlobStorageRepositoryTests
{
    [Fact]
    public void Ctor_WithConnectionString_DoesNotThrow()
    {
        var repo = new AzureBlobStorageRepository(
            connectionString: "UseDevelopmentStorage=true",
            containerName: "test",
            defaultTtl: TimeSpan.FromMinutes(5));
        Assert.NotNull(repo);
    }

    [Fact]
    public void Ctor_WithUriAndNoCredential_DoesNotThrow()
    {
        var repo = new AzureBlobStorageRepository(
            containerUri: new Uri("https://example.com/container"),
            credential: null,
            defaultTtl: TimeSpan.FromMinutes(5));
        Assert.NotNull(repo);
    }

    [Fact]
    public void Ctor_WithContainer_DoesNotThrow()
    {
        var container = new Mock<IBlobContainerClient>();
        var repo = new AzureBlobStorageRepository(container.Object, TimeSpan.FromMinutes(5));
        Assert.NotNull(repo);
    }

    [Fact]
    public async Task UploadExcelAsync_WhenCannotGenerateSas_UsesBlobUri()
    {
        // Arrange
        var container = new Mock<IBlobContainerClient>();
        var blob = new Mock<IBlobClient>();

        container.Setup(c => c.Name).Returns("test");
        container
            .Setup(c => c.CreateIfNotExistsAsync(PublicAccessType.None, It.IsAny<CancellationToken>()))
            .Returns(Task.CompletedTask);
        container
            .Setup(c => c.GetBlobClient(It.IsAny<string>()))
            .Returns(blob.Object);

        blob.Setup(b => b.CanGenerateSasUri).Returns(false); // force else branch
        var blobUri = new Uri("https://example.com/blob");
        blob.Setup(b => b.Uri).Returns(blobUri);
        blob
            .Setup(b => b.UploadAsync(It.IsAny<Stream>(), It.IsAny<BlobUploadOptions>(), It.IsAny<CancellationToken>()))
            .Returns(Task.CompletedTask);

        var repo = new AzureBlobStorageRepository(container.Object, TimeSpan.FromMinutes(10));
        await using var ms = new MemoryStream(new byte[] { 1, 2, 3, 4 });

        // Act
        var response = await repo.UploadExcelAsync(ms, "file.xlsx", sasTtl: null, ct: default);

        // Assert
        Assert.True(response.IsSuccess);
        Assert.NotNull(response.Data);
        Assert.Equal(blobUri, response.Data!.SasReadUri);
    }

    [Fact]
    public async Task UploadExcelAsync_OnException_ReturnsFailure()
    {
        // Arrange
        var container = new Mock<IBlobContainerClient>();
        container
            .Setup(c => c.CreateIfNotExistsAsync(PublicAccessType.None, It.IsAny<CancellationToken>()))
            .ThrowsAsync(new InvalidOperationException("boom"));

        var repo = new AzureBlobStorageRepository(container.Object, TimeSpan.FromMinutes(10));
        await using var ms = new MemoryStream(new byte[] { 1, 2, 3 });

        // Act
        var response = await repo.UploadExcelAsync(ms, "file.xlsx", sasTtl: null, ct: default);

        // Assert
        Assert.False(response.IsSuccess);
        Assert.NotNull(response.ErrorMessage);
    }
}

