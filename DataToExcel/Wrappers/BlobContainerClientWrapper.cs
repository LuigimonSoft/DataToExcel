using System.Diagnostics.CodeAnalysis;
using Azure.Storage.Blobs;
using Azure.Storage.Blobs.Models;
using DataToExcel.Wrappers.Interfaces;

namespace DataToExcel.Wrappers;

[ExcludeFromCodeCoverage]
public class BlobContainerClientWrapper : IBlobContainerClient
{
    private readonly BlobContainerClient _inner;

    public BlobContainerClientWrapper(BlobContainerClient inner) => _inner = inner;

    public string Name => _inner.Name;

    public async Task CreateIfNotExistsAsync(PublicAccessType accessType, CancellationToken ct)
        => await _inner.CreateIfNotExistsAsync(accessType, cancellationToken: ct);

    public IBlobClient GetBlobClient(string blobName) => new BlobClientWrapper(_inner.GetBlobClient(blobName));
}
