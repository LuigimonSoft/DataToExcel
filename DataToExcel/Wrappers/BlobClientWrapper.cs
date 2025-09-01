using Azure.Storage.Blobs;
using Azure.Storage.Blobs.Models;
using Azure.Storage.Sas;
using DataToExcel.Wrappers.Interfaces;

namespace DataToExcel.Wrappers;

public class BlobClientWrapper : IBlobClient
{
    private readonly BlobClient _inner;

    public BlobClientWrapper(BlobClient inner) => _inner = inner;

    public bool CanGenerateSasUri => _inner.CanGenerateSasUri;

    public Uri Uri => _inner.Uri;

    public async Task UploadAsync(Stream content, BlobUploadOptions options, CancellationToken ct)
        => await _inner.UploadAsync(content, options, ct);

    public Uri GenerateSasUri(BlobSasBuilder builder) => _inner.GenerateSasUri(builder);
}
