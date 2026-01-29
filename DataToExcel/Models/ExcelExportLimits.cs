namespace DataToExcel.Models;

public static class ExcelExportLimits
{
    public const int MaxRowsPerSheet = 1_048_576;
    public const int HeaderRowCount = 1;
    public const int MaxDataRowsPerSheet = MaxRowsPerSheet - HeaderRowCount;
}
