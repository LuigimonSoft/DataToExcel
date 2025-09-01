using DataToExcel.Models;

namespace DataToExcel.Repositories.Interfaces;

public interface IBlobStorageRepository
{
    Task<RepositoryResponse<BlobUploadResult>> UploadExcelAsync(Stream excelStream, string blobName, TimeSpan? sasTtl, CancellationToken ct = default);
}
