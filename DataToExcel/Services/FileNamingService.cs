using DataToExcel.Models;
using DataToExcel.Services.Interfaces;

namespace DataToExcel.Services;

public class FileNamingService : IFileNamingService
{
    public ServiceResponse<string> ComposeExcelFileName(string baseName, DateTime dataDateUtc, DateTime creationUtc)
    {
        try
        {
            var invalid = Path.GetInvalidFileNameChars();
            var sanitized = new string(baseName.Select(c => invalid.Contains(c) ? '_' : c).ToArray());
            var name = $"{sanitized}_{dataDateUtc:yyyyMMdd}_{creationUtc:yyyyMMdd_HHmmss}.xlsx";
            return new ServiceResponse<string>(name) { IsSuccess = true };
        }
        catch (Exception ex)
        {
            return new ServiceResponse<string> { IsSuccess = false, ErrorMessage = ex.Message };
        }
    }
}
