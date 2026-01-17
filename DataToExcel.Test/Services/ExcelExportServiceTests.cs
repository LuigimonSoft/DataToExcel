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
        IEnumerable<IDataRecord> Records()
        {
            var reader = table.CreateDataReader();
            while (reader.Read()) yield return reader;
        }
        var records = Records();

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
    public async Task GivenAsyncRecordsWhenExportAsyncThenHeaderShouldBeWritten()
    {
        var table = new DataTable();
        table.Columns.Add("Name", typeof(string));
        table.Rows.Add("Alice");
        var records = ToAsyncEnumerable(table);

        var columns = new List<ColumnDefinition>
        {
            new("Name","Name", ColumnDataType.String)
        };
        var service = new ExcelExportService(new ExcelStyleProvider());
        using var ms = new MemoryStream();

        var response = await service.ExportAsync(records, columns, ms, new ExcelExportOptions());

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
        IEnumerable<IDataRecord> Records2()
        {
            var reader = table.CreateDataReader();
            while (reader.Read()) yield return reader;
        }
        var records = Records2();

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

    [Fact]
    public async Task GivenAsyncRecordsWithMultipleColumnsWhenExportAsyncThenValuesShouldBeInCorrectCells()
    {
        var table = new DataTable();
        table.Columns.Add("Name", typeof(string));
        table.Columns.Add("Age", typeof(int));
        table.Rows.Add("Alice", 30);
        var records = ToAsyncEnumerable(table);

        var columns = new List<ColumnDefinition>
        {
            new("Name","Name", ColumnDataType.String),
            new("Age","Age", ColumnDataType.Number)
        };
        var service = new ExcelExportService(new ExcelStyleProvider());
        using var ms = new MemoryStream();

        var response = await service.ExportAsync(records, columns, ms, new ExcelExportOptions());

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

    [Fact]
    public async Task GivenHiddenColumnWhenExportAsyncThenColumnShouldBeHidden()
    {
        var table = new DataTable();
        table.Columns.Add("Name", typeof(string));
        table.Rows.Add("Alice");
        IEnumerable<IDataRecord> Records3()
        {
            var reader = table.CreateDataReader();
            while (reader.Read()) yield return reader;
        }
        var records = Records3();

        var columns = new List<ColumnDefinition>
        {
            new("Name","Name", ColumnDataType.String, Hidden: true)
        };
        var service = new ExcelExportService(new ExcelStyleProvider());
        using var ms = new MemoryStream();

        var response = await service.ExportAsync(records, columns, ms, new ExcelExportOptions());

        Assert.True(response.IsSuccess);
        ms.Position = 0;
        using var doc = SpreadsheetDocument.Open(ms, false);
        var sheet = doc.WorkbookPart!.WorksheetParts.First().Worksheet;
        var column = sheet.GetFirstChild<Columns>()!.Elements<Column>().First();
        Assert.True(column.Hidden!.Value);
    }

    [Fact]
    public async Task GivenHiddenColumnWhenExportAsyncAsyncRecordsThenColumnShouldBeHidden()
    {
        var table = new DataTable();
        table.Columns.Add("Name", typeof(string));
        table.Rows.Add("Alice");
        var records = ToAsyncEnumerable(table);

        var columns = new List<ColumnDefinition>
        {
            new("Name","Name", ColumnDataType.String, Hidden: true)
        };
        var service = new ExcelExportService(new ExcelStyleProvider());
        using var ms = new MemoryStream();

        var response = await service.ExportAsync(records, columns, ms, new ExcelExportOptions());

        Assert.True(response.IsSuccess);
        ms.Position = 0;
        using var doc = SpreadsheetDocument.Open(ms, false);
        var sheet = doc.WorkbookPart!.WorksheetParts.First().Worksheet;
        var column = sheet.GetFirstChild<Columns>()!.Elements<Column>().First();
        Assert.True(column.Hidden!.Value);
    }

    [Fact]
    public async Task GivenGroupedColumnWhenExportAsyncThenRowsShouldBeGrouped()
    {
        var table = new DataTable();
        table.Columns.Add("Category", typeof(string));
        table.Columns.Add("Amount", typeof(int));
        table.Rows.Add("A", 1);
        table.Rows.Add("A", 2);
        table.Rows.Add("B", 3);
        IEnumerable<IDataRecord> Records4()
        {
            var reader = table.CreateDataReader();
            while (reader.Read()) yield return reader;
        }
        var records = Records4();

        var columns = new List<ColumnDefinition>
        {
            new("Category","Category", ColumnDataType.String, Group: true),
            new("Amount","Amount", ColumnDataType.Number)
        };
        var service = new ExcelExportService(new ExcelStyleProvider());
        using var ms = new MemoryStream();

        var response = await service.ExportAsync(records, columns, ms, new ExcelExportOptions());

        Assert.True(response.IsSuccess);
        ms.Position = 0;
        using var doc = SpreadsheetDocument.Open(ms, false);
        var sheet = doc.WorkbookPart!.WorksheetParts.First().Worksheet;
        var rows = sheet.GetFirstChild<SheetData>()!.Elements<Row>().ToList();

        Assert.Equal("A", rows[1].Elements<Cell>().First().InnerText);
        Assert.Null(rows[1].OutlineLevel);
        Assert.Equal((byte)1, rows[2].OutlineLevel!.Value);
        Assert.Equal("", rows[2].Elements<Cell>().First().InnerText);
        Assert.Equal("B", rows[3].Elements<Cell>().First().InnerText);
        var sheetFormat = sheet.Elements<SheetFormatProperties>().FirstOrDefault();
        Assert.NotNull(sheetFormat);
        Assert.Equal((byte)1, sheetFormat!.OutlineLevelRow!.Value);
    }

    [Fact]
    public async Task GivenGroupedColumnWhenExportAsyncAsyncRecordsThenRowsShouldBeGrouped()
    {
        var table = new DataTable();
        table.Columns.Add("Category", typeof(string));
        table.Columns.Add("Amount", typeof(int));
        table.Rows.Add("A", 1);
        table.Rows.Add("A", 2);
        table.Rows.Add("B", 3);
        var records = ToAsyncEnumerable(table);

        var columns = new List<ColumnDefinition>
        {
            new("Category","Category", ColumnDataType.String, Group: true),
            new("Amount","Amount", ColumnDataType.Number)
        };
        var service = new ExcelExportService(new ExcelStyleProvider());
        using var ms = new MemoryStream();

        var response = await service.ExportAsync(records, columns, ms, new ExcelExportOptions());

        Assert.True(response.IsSuccess);
        ms.Position = 0;
        using var doc = SpreadsheetDocument.Open(ms, false);
        var sheet = doc.WorkbookPart!.WorksheetParts.First().Worksheet;
        var rows = sheet.GetFirstChild<SheetData>()!.Elements<Row>().ToList();

        Assert.Equal("A", rows[1].Elements<Cell>().First().InnerText);
        Assert.Null(rows[1].OutlineLevel);
        Assert.Equal((byte)1, rows[2].OutlineLevel!.Value);
        Assert.Equal("", rows[2].Elements<Cell>().First().InnerText);
        Assert.Equal("B", rows[3].Elements<Cell>().First().InnerText);
        var sheetFormat = sheet.Elements<SheetFormatProperties>().FirstOrDefault();
        Assert.NotNull(sheetFormat);
        Assert.Equal((byte)1, sheetFormat!.OutlineLevelRow!.Value);
    }

    private static async IAsyncEnumerable<IDataRecord> ToAsyncEnumerable(DataTable table)
    {
        using var reader = table.CreateDataReader();
        while (reader.Read())
        {
            await Task.Yield();
            yield return reader;
        }
    }
}
