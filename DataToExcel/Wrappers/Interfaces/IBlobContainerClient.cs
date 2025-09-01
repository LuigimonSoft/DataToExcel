using Azure.Storage.Blobs.Models;

namespace DataToExcel.Wrappers.Interfaces;

public interface IBlobContainerClient
{
    string Name { get; }
    Task CreateIfNotExistsAsync(PublicAccessType accessType, CancellationToken ct);
    IBlobClient GetBlobClient(string blobName);
}
