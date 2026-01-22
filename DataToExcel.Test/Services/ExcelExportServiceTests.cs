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
        var table = BuildGroupedTable();
        var records = ToRecords(table);

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
        Assert.Equal("B", rows[6].Elements<Cell>().First().InnerText);
        Assert.Equal("C", rows[11].Elements<Cell>().First().InnerText);
        Assert.Equal("D", rows[16].Elements<Cell>().First().InnerText);
        Assert.Equal(4, rows.Skip(1).Count(r => r.OutlineLevel is null));
        Assert.Equal(16, rows.Skip(1).Count(r => r.OutlineLevel?.Value == 1));
        Assert.All(rows.Skip(1).Where(r => r.OutlineLevel?.Value == 1), r =>
        {
            Assert.Equal(string.Empty, r.Elements<Cell>().First().InnerText);
        });
        var sheetFormat = sheet.Elements<SheetFormatProperties>().FirstOrDefault();
        Assert.NotNull(sheetFormat);
        Assert.Equal((byte)1, sheetFormat!.OutlineLevelRow!.Value);
    }

    [Fact]
    public async Task GivenGroupedColumnWhenExportAsyncAsyncRecordsThenRowsShouldBeGrouped()
    {
        var table = BuildGroupedTable();
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
        Assert.Equal("B", rows[6].Elements<Cell>().First().InnerText);
        Assert.Equal("C", rows[11].Elements<Cell>().First().InnerText);
        Assert.Equal("D", rows[16].Elements<Cell>().First().InnerText);
        Assert.Equal(4, rows.Skip(1).Count(r => r.OutlineLevel is null));
        Assert.Equal(16, rows.Skip(1).Count(r => r.OutlineLevel?.Value == 1));
        Assert.All(rows.Skip(1).Where(r => r.OutlineLevel?.Value == 1), r =>
        {
            Assert.Equal(string.Empty, r.Elements<Cell>().First().InnerText);
        });
        var sheetFormat = sheet.Elements<SheetFormatProperties>().FirstOrDefault();
        Assert.NotNull(sheetFormat);
        Assert.Equal((byte)1, sheetFormat!.OutlineLevelRow!.Value);
    }

    [Fact]
    public async Task GivenGroupedMiddleColumnWhenExportAsyncThenRowsShouldBeGrouped()
    {
        var table = BuildGroupedTable(withItem: true);
        var records = ToRecords(table);

        var columns = new List<ColumnDefinition>
        {
            new("Item","Item", ColumnDataType.String),
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

        var firstGroupRowCells = rows[1].Elements<Cell>().ToList();
        Assert.Null(rows[1].OutlineLevel);
        Assert.Equal("Item A-1", firstGroupRowCells[0].InnerText);
        Assert.Equal("A", firstGroupRowCells[1].InnerText);
        Assert.Equal("10", firstGroupRowCells[2].CellValue!.Text);

        var groupedDetailCells = rows[2].Elements<Cell>().ToList();
        Assert.Equal((byte)1, rows[2].OutlineLevel!.Value);
        Assert.Equal("Item A-2", groupedDetailCells[0].InnerText);
        Assert.Equal(string.Empty, groupedDetailCells[1].InnerText);
        Assert.Equal("20", groupedDetailCells[2].CellValue!.Text);

        var secondGroupRowCells = rows[6].Elements<Cell>().ToList();
        Assert.Null(rows[6].OutlineLevel);
        Assert.Equal("Item B-1", secondGroupRowCells[0].InnerText);
        Assert.Equal("B", secondGroupRowCells[1].InnerText);
        Assert.Equal("60", secondGroupRowCells[2].CellValue!.Text);
    }

    [Fact]
    public async Task GivenGroupedMiddleColumnWhenExportAsyncAsyncRecordsThenRowsShouldBeGrouped()
    {
        var table = BuildGroupedTable(withItem: true);
        var records = ToAsyncEnumerable(table);

        var columns = new List<ColumnDefinition>
        {
            new("Item","Item", ColumnDataType.String),
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

        var firstGroupRowCells = rows[1].Elements<Cell>().ToList();
        Assert.Null(rows[1].OutlineLevel);
        Assert.Equal("Item A-1", firstGroupRowCells[0].InnerText);
        Assert.Equal("A", firstGroupRowCells[1].InnerText);
        Assert.Equal("10", firstGroupRowCells[2].CellValue!.Text);

        var groupedDetailCells = rows[2].Elements<Cell>().ToList();
        Assert.Equal((byte)1, rows[2].OutlineLevel!.Value);
        Assert.Equal("Item A-2", groupedDetailCells[0].InnerText);
        Assert.Equal(string.Empty, groupedDetailCells[1].InnerText);
        Assert.Equal("20", groupedDetailCells[2].CellValue!.Text);

        var secondGroupRowCells = rows[6].Elements<Cell>().ToList();
        Assert.Null(rows[6].OutlineLevel);
        Assert.Equal("Item B-1", secondGroupRowCells[0].InnerText);
        Assert.Equal("B", secondGroupRowCells[1].InnerText);
        Assert.Equal("60", secondGroupRowCells[2].CellValue!.Text);
    }

    [Fact]
    public async Task GivenForwardOnlyAsyncRecordsWhenExportAsyncThenRowsShouldBeGrouped()
    {
        var table = BuildGroupedTable();
        var records = new ForwardOnlyAsyncRecords(table);

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

        Assert.Equal(20, rows.Skip(1).Count());
        Assert.Equal(4, rows.Skip(1).Count(r => r.OutlineLevel is null));
        Assert.Equal(16, rows.Skip(1).Count(r => r.OutlineLevel?.Value == 1));
        Assert.All(rows.Skip(1).Where(r => r.OutlineLevel?.Value == 1), r =>
        {
            Assert.Equal(string.Empty, r.Elements<Cell>().First().InnerText);
        });
    }

    private static DataTable BuildGroupedTable(bool withItem = false)
    {
        var table = new DataTable();
        if (withItem)
        {
            table.Columns.Add("Item", typeof(string));
        }
        table.Columns.Add("Category", typeof(string));
        table.Columns.Add("Amount", typeof(int));

        var groups = new[] { "A", "B", "C", "D" };
        foreach (var group in groups)
        {
            for (var i = 1; i <= 5; i++)
            {
                var amount = ((Array.IndexOf(groups, group) * 5) + i) * 10;
                if (withItem)
                {
                    table.Rows.Add($"Item {group}-{i}", group, amount);
                }
                else
                {
                    table.Rows.Add(group, amount);
                }
            }
        }

        return table;
    }

    private static IEnumerable<IDataRecord> ToRecords(DataTable table)
    {
        var reader = table.CreateDataReader();
        while (reader.Read())
        {
            yield return reader;
        }
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

    private sealed class ForwardOnlyAsyncRecords : IAsyncEnumerable<IDataRecord>, IAsyncEnumerator<IDataRecord>
    {
        private readonly DataTable _table;
        private DataTableReader? _reader;
        private bool _started;

        public ForwardOnlyAsyncRecords(DataTable table)
        {
            _table = table;
        }

        public IDataRecord Current => _reader ?? throw new InvalidOperationException("Enumerator not started.");

        public IAsyncEnumerator<IDataRecord> GetAsyncEnumerator(CancellationToken cancellationToken = default)
        {
            if (_started)
            {
                throw new InvalidOperationException("This enumerator can only be iterated once.");
            }

            _started = true;
            _reader = _table.CreateDataReader();
            return this;
        }

        public ValueTask DisposeAsync()
        {
            _reader?.Dispose();
            _reader = null;
            return ValueTask.CompletedTask;
        }

        public ValueTask<bool> MoveNextAsync()
        {
            if (_reader is null)
            {
                throw new InvalidOperationException("Enumerator not started.");
            }

            return new ValueTask<bool>(_reader.Read());
        }
    }
}
