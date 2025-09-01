namespace DataToExcel.Models;

public record BlobUploadResult(
    string Container,
    string BlobName,
    Uri BlobUri,
    Uri SasReadUri,
    long SizeBytes);
