using Azure.Storage.Blobs.Models;
using Azure.Storage.Sas;

namespace DataToExcel.Wrappers.Interfaces;

public interface IBlobClient
{
    bool CanGenerateSasUri { get; }
    Uri Uri { get; }
    Task UploadAsync(Stream content, BlobUploadOptions options, CancellationToken ct);
    Uri GenerateSasUri(BlobSasBuilder builder);
}
