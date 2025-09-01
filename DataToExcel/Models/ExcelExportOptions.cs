using System.Globalization;

namespace DataToExcel.Models;

public class ExcelExportOptions
{
    public string SheetName { get; set; } = "Sheet1";
    public CultureInfo Culture { get; set; } = CultureInfo.InvariantCulture;
    public bool FreezeHeader { get; set; } = true;
    public bool AutoFilter { get; set; } = true;
    public DateTime? DataDateUtc { get; set; } = DateTime.UtcNow.Date;
}
