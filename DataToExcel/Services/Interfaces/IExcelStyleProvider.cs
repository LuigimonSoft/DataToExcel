using DocumentFormat.OpenXml.Spreadsheet;
using DataToExcel.Models;

namespace DataToExcel.Services.Interfaces;

public interface IExcelStyleProvider
{
    ServiceResponse<Stylesheet> BuildStylesheet(ExcelExportOptions options, out IReadOnlyDictionary<PredefinedStyle, uint> styleIndexMap);
}
