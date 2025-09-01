using DataToExcel.Models;

namespace DataToExcel.Services.Interfaces;

public interface IFileNamingService
{
    ServiceResponse<string> ComposeExcelFileName(string baseName, DateTime dataDateUtc, DateTime creationUtc);
}
