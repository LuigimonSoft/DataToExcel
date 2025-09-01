using Azure;
using Azure.Core;
using Azure.Storage;
using Azure.Storage.Blobs;
using Azure.Storage.Blobs.Models;
using Azure.Storage.Sas;
using DataToExcel.Models;
using DataToExcel.Repositories.Interfaces;
using DataToExcel.Wrappers;
using DataToExcel.Wrappers.Interfaces;

namespace DataToExcel.Repositories;

public class AzureBlobStorageRepository : IBlobStorageRepository
{
    private readonly IBlobContainerClient _container;
    private readonly TimeSpan _defaultTtl;

    public AzureBlobStorageRepository(string connectionString, string containerName, TimeSpan defaultTtl)
        : this(new BlobContainerClientWrapper(new BlobContainerClient(connectionString, containerName)), defaultTtl)
    {
    }

    public AzureBlobStorageRepository(Uri containerUri, TokenCredential? credential, TimeSpan defaultTtl)
        : this(
            new BlobContainerClientWrapper(
                credential is null
                    ? new BlobContainerClient(containerUri)
                    : new BlobContainerClient(containerUri, credential)),
            defaultTtl)
    {
    }

    public AzureBlobStorageRepository(IBlobContainerClient container, TimeSpan defaultTtl)
    {
        _container = container;
        _defaultTtl = defaultTtl;
    }

    public async Task<RepositoryResponse<BlobUploadResult>> UploadExcelAsync(Stream excelStream, string blobName, TimeSpan? sasTtl, CancellationToken ct = default)
    {
        try
        {
            await _container.CreateIfNotExistsAsync(PublicAccessType.None, ct);
            var blobClient = _container.GetBlobClient(blobName);
            excelStream.Position = 0;
            var options = new BlobUploadOptions
            {
                HttpHeaders = new BlobHttpHeaders { ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" },
                TransferOptions = new StorageTransferOptions
                {
                    InitialTransferSize = 8 * 1024 * 1024,
                    MaximumTransferSize = 8 * 1024 * 1024,
                    MaximumConcurrency = Environment.ProcessorCount
                }
            };
            await blobClient.UploadAsync(excelStream, options, ct);

            var expiry = DateTimeOffset.UtcNow.Add(sasTtl ?? _defaultTtl);
            Uri sasUri;
            if (blobClient.CanGenerateSasUri)
            {
                var builder = new BlobSasBuilder(BlobSasPermissions.Read, expiry)
                {
                    BlobContainerName = _container.Name,
                    BlobName = blobName
                };
                sasUri = blobClient.GenerateSasUri(builder);
            }
            else
            {
                sasUri = blobClient.Uri; // fallback without SAS
            }

            var result = new BlobUploadResult(_container.Name, blobName, blobClient.Uri, sasUri, excelStream.Length);
            return new RepositoryResponse<BlobUploadResult>(result) { IsSuccess = true };
        }
        catch (Exception ex)
        {
            return new RepositoryResponse<BlobUploadResult> { IsSuccess = false, ErrorMessage = ex.Message };
        }
    }
}
