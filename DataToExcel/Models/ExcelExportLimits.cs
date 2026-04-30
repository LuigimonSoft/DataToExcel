namespace DataToExcel.Models;

public static class ExcelExportLimits
{
    public const int MaxRowsPerSheet = 1_048_576;
    public const int HeaderRowCount = 1;
    public const int MaxDataRowsPerSheet = MaxRowsPerSheet - HeaderRowCount;

    // System.IO.Packaging/OpenXML can fail before Excel's row limit when a
    // worksheet part grows too large. Keep a margin under 2 GB because the
    // exact serialized XML size varies with values, styles, and row numbers.
    public const long MaxWorksheetPartBytes = 1_700_000_000;
}
