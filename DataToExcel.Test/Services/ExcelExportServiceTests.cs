using System.Data;
using DataToExcel.Models;
using DataToExcel.Services;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Linq;
using Xunit;

namespace DataToExcel.Test.Services;

public class ExcelExportServiceTests
{
    [Fact]
    public async Task GivenRecordsWhenExportAsyncThenHeaderShouldBeWritten()
    {
        // Given
        var table = new DataTable();
        table.Columns.Add("Name", typeof(string));
        table.Rows.Add("Alice");
        var reader = table.CreateDataReader();
        var records = new List<IDataRecord>();
        while (reader.Read()) records.Add(reader);

        var columns = new List<ColumnDefinition>
        {
            new("Name","Name", ColumnDataType.String)
        };
        var service = new ExcelExportService(new ExcelStyleProvider());
        using var ms = new MemoryStream();

        // When
        var response = await service.ExportAsync(records, columns, ms, new ExcelExportOptions());

        // Then
        Assert.True(response.IsSuccess);
        ms.Position = 0;
        using var doc = SpreadsheetDocument.Open(ms, false);
        var sheet = doc.WorkbookPart!.WorksheetParts.First().Worksheet;
        var header = sheet.GetFirstChild<SheetData>()!.Elements<Row>().First().Elements<Cell>().First().CellValue!.Text;
        Assert.Equal("Name", header);
    }

    [Fact]
    public async Task GivenRecordsWithMultipleColumnsWhenExportAsyncThenValuesShouldBeInCorrectCells()
    {
        // Given
        var table = new DataTable();
        table.Columns.Add("Name", typeof(string));
        table.Columns.Add("Age", typeof(int));
        table.Rows.Add("Alice", 30);
        var reader = table.CreateDataReader();
        var records = new List<IDataRecord>();
        while (reader.Read()) records.Add(reader);

        var columns = new List<ColumnDefinition>
        {
            new("Name","Name", ColumnDataType.String),
            new("Age","Age", ColumnDataType.Number)
        };
        var service = new ExcelExportService(new ExcelStyleProvider());
        using var ms = new MemoryStream();

        // When
        var response = await service.ExportAsync(records, columns, ms, new ExcelExportOptions());

        // Then
        Assert.True(response.IsSuccess);
        ms.Position = 0;
        using var doc = SpreadsheetDocument.Open(ms, false);
        var sheet = doc.WorkbookPart!.WorksheetParts.First().Worksheet;
        var rows = sheet.GetFirstChild<SheetData>()!.Elements<Row>().ToList();
        var dataRow = rows[1];
        var cells = dataRow.Elements<Cell>().ToList();
        Assert.Equal("Alice", cells[0].InnerText);
        Assert.Equal("30", cells[1].CellValue!.Text);
    }
}
