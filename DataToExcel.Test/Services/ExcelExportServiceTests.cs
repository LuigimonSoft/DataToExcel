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
    public async Task GivenSplitIntoMultipleSheetsWhenExportAsyncThenOptionsAreHonored()
    {
        var table = new DataTable();
        table.Columns.Add("Name", typeof(string));
        table.Rows.Add("Alice");
        var records = ToAsyncEnumerable(table);

        var columns = new List<ColumnDefinition>
        {
            new("Name", "Name", ColumnDataType.String, Width: 20, Hidden: true)
        };
        var service = new ExcelExportService(new ExcelStyleProvider());
        using var ms = new MemoryStream();

        var response = await service.ExportAsync(records, columns, ms, new ExcelExportOptions
        {
            SplitIntoMultipleSheets = true,
            FreezeHeader = false,
            AutoFilter = false
        });

        Assert.True(response.IsSuccess);
        ms.Position = 0;
        using var doc = SpreadsheetDocument.Open(ms, false);
        var sheet = doc.WorkbookPart!.WorksheetParts.First().Worksheet;
        Assert.Null(sheet.SheetViews);
        Assert.Null(sheet.Elements<AutoFilter>().FirstOrDefault());
        Assert.NotNull(sheet.Elements<Columns>().FirstOrDefault());
    }

    [Fact]
    public void GivenLongSheetNameWhenComposeSheetNameThenTrimsAndSuffixes()
    {
        var method = typeof(ExcelExportService).GetMethod("ComposeSheetName",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        Assert.NotNull(method);

        var longName = new string('A', 40);
        var result = method!.Invoke(null, new object?[] { longName, 2 }) as string;

        Assert.NotNull(result);
        Assert.EndsWith(" (2)", result, StringComparison.Ordinal);
        Assert.True(result!.Length <= 31);
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
    public async Task GivenFirstGroupedValueIsNullWhenExportAsyncThenFirstRowStartsGroup()
    {
        var table = new DataTable();
        table.Columns.Add("Category", typeof(string));
        table.Columns.Add("Amount", typeof(int));
        table.Rows.Add(DBNull.Value, 10);
        table.Rows.Add(DBNull.Value, 20);
        table.Rows.Add("A", 30);
        var records = ToAsyncEnumerable(table);

        var columns = new List<ColumnDefinition>
        {
            new("Category", "Category", ColumnDataType.String, Group: true),
            new("Amount", "Amount", ColumnDataType.Number)
        };
        var service = new ExcelExportService(new ExcelStyleProvider());
        using var ms = new MemoryStream();

        var response = await service.ExportAsync(records, columns, ms, new ExcelExportOptions());

        Assert.True(response.IsSuccess);
        ms.Position = 0;
        using var doc = SpreadsheetDocument.Open(ms, false);
        var rows = doc.WorkbookPart!.WorksheetParts.First().Worksheet.GetFirstChild<SheetData>()!.Elements<Row>().ToList();

        Assert.Null(rows[1].OutlineLevel);
        Assert.Equal(string.Empty, rows[1].Elements<Cell>().First().InnerText);
        Assert.Equal("10", rows[1].Elements<Cell>().Last().CellValue!.Text);
        Assert.Equal((byte)1, rows[2].OutlineLevel!.Value);
        Assert.Null(rows[3].OutlineLevel);
        Assert.Equal("A", rows[3].Elements<Cell>().First().InnerText);
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

    [Fact]
    public async Task GivenForwardOnlyAsyncRecordsWithColumnOrderMismatchWhenExportAsyncThenCellsFollowColumnDefinition()
    {
        var table = new DataTable();
        table.Columns.Add("Name", typeof(string));
        table.Columns.Add("Age", typeof(int));
        table.Rows.Add("Alice", 30);
        var queryBuilder = new DataTableProjectionBuilder(table)
            .Select("Age", "Name");
        var executor = new DataTableProjectionExecutor();
        var records = executor.ExecuteAsync(queryBuilder);

        var columns = new List<ColumnDefinition>
        {
            new("Age","Age", ColumnDataType.Number),
            new("Name","Name", ColumnDataType.String)
        };
        var service = new ExcelExportService(new ExcelStyleProvider());
        using var ms = new MemoryStream();

        var response = await service.ExportAsync(records, columns, ms, new ExcelExportOptions());

        Assert.True(response.IsSuccess);
        ms.Position = 0;
        using var doc = SpreadsheetDocument.Open(ms, false);
        var sheet = doc.WorkbookPart!.WorksheetParts.First().Worksheet;
        var rows = sheet.GetFirstChild<SheetData>()!.Elements<Row>().ToList();
        var headerCells = rows[0].Elements<Cell>().ToList();
        var dataCells = rows[1].Elements<Cell>().ToList();

        Assert.Equal("Age", headerCells[0].CellValue!.Text);
        Assert.Equal("Name", headerCells[1].CellValue!.Text);
        Assert.Equal("30", dataCells[0].CellValue!.Text);
        Assert.Equal("Alice", dataCells[1].InnerText);
    }


    [Fact]
    public async Task GivenCaseMismatchedFieldNamesWhenExportAsyncThenValuesAreResolvedCaseInsensitive()
    {
        var table = new DataTable();
        table.Columns.Add("Name", typeof(string));
        table.Columns.Add("Age", typeof(int));
        table.Rows.Add("Alice", 30);
        var records = ToAsyncEnumerable(table);

        var columns = new List<ColumnDefinition>
        {
            new("name","name", ColumnDataType.String),
            new("age","age", ColumnDataType.Number)
        };

        var service = new ExcelExportService(new ExcelStyleProvider());
        using var ms = new MemoryStream();
        var response = await service.ExportAsync(records, columns, ms, new ExcelExportOptions());

        Assert.True(response.IsSuccess);
        ms.Position = 0;
        using var doc = SpreadsheetDocument.Open(ms, false);
        var rows = doc.WorkbookPart!.WorksheetParts.First().Worksheet.GetFirstChild<SheetData>()!.Elements<Row>().ToList();
        var dataCells = rows[1].Elements<Cell>().ToList();
        Assert.Equal("Alice", dataCells[0].InnerText);
        Assert.Equal("30", dataCells[1].CellValue!.Text);
    }

    [Fact]
    public async Task GivenLongSheetNameWithoutSplitWhenExportAsyncThenSheetNameIsTrimmedTo31Chars()
    {
        var table = new DataTable();
        table.Columns.Add("Name", typeof(string));
        table.Rows.Add("Alice");

        var columns = new List<ColumnDefinition> { new("Name", "Name", ColumnDataType.String) };
        var options = new ExcelExportOptions { SheetName = new string('X', 40), SplitIntoMultipleSheets = false };
        var service = new ExcelExportService(new ExcelStyleProvider());
        using var ms = new MemoryStream();

        var response = await service.ExportAsync(ToAsyncEnumerable(table), columns, ms, options);

        Assert.True(response.IsSuccess);
        ms.Position = 0;
        using var doc = SpreadsheetDocument.Open(ms, false);
        var sheet = doc.WorkbookPart!.Workbook.Sheets!.Elements<Sheet>().First();
        Assert.Equal(31, sheet.Name!.Value!.Length);
        Assert.Equal(new string('X', 31), sheet.Name!.Value);
    }

    [Fact]
    public async Task GivenNonSeekableOutputWhenExportAsyncThenWorkbookIsWrittenUsingStaging()
    {
        var table = new DataTable();
        table.Columns.Add("Name", typeof(string));
        table.Rows.Add("Alice");

        var service = new ExcelExportService(new ExcelStyleProvider());
        var columns = new List<ColumnDefinition> { new("Name", "Name", ColumnDataType.String) };
        using var nonSeekable = new NonSeekableWriteOnlyStream();

        var response = await service.ExportAsync(ToAsyncEnumerable(table), columns, nonSeekable, new ExcelExportOptions());

        Assert.True(response.IsSuccess);
        var bytes = nonSeekable.ToArray();
        using var resultStream = new MemoryStream(bytes);
        using var doc = SpreadsheetDocument.Open(resultStream, false);
        var dataRows = doc.WorkbookPart!.WorksheetParts.First().Worksheet.GetFirstChild<SheetData>()!.Elements<Row>().ToList();
        Assert.Equal("Alice", dataRows[1].Elements<Cell>().First().InnerText);
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

    private sealed class DataTableProjectionBuilder
    {
        private readonly DataTable _table;
        private IReadOnlyList<string> _selectedColumns = Array.Empty<string>();

        public DataTableProjectionBuilder(DataTable table)
        {
            _table = table;
        }

        public DataTableProjectionBuilder Select(params string[] columns)
        {
            _selectedColumns = columns;
            return this;
        }

        public DataTable Build()
        {
            if (_selectedColumns.Count == 0)
            {
                return _table;
            }

            var projection = new DataTable();
            foreach (var column in _selectedColumns)
            {
                projection.Columns.Add(column, _table.Columns[column]!.DataType);
            }

            foreach (DataRow row in _table.Rows)
            {
                var values = _selectedColumns.Select(col => row[col]).ToArray();
                projection.Rows.Add(values);
            }

            return projection;
        }
    }

    private sealed class DataTableProjectionExecutor
    {
        public IAsyncEnumerable<IDataRecord> ExecuteAsync(DataTableProjectionBuilder builder)
        {
            var table = builder.Build();
            return new ForwardOnlyAsyncRecords(table);
        }
    }

    private sealed class NonSeekableWriteOnlyStream : Stream
    {
        private readonly MemoryStream _inner = new();

        public override bool CanRead => false;
        public override bool CanSeek => false;
        public override bool CanWrite => true;
        public override long Length => _inner.Length;
        public override long Position
        {
            get => throw new NotSupportedException();
            set => throw new NotSupportedException();
        }

        public override void Flush() => _inner.Flush();
        public override Task FlushAsync(CancellationToken cancellationToken) => _inner.FlushAsync(cancellationToken);
        public override int Read(byte[] buffer, int offset, int count) => throw new NotSupportedException();
        public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
        public override void SetLength(long value) => _inner.SetLength(value);
        public override void Write(byte[] buffer, int offset, int count) => _inner.Write(buffer, offset, count);
        public override ValueTask WriteAsync(ReadOnlyMemory<byte> buffer, CancellationToken cancellationToken = default)
            => _inner.WriteAsync(buffer, cancellationToken);

        public byte[] ToArray() => _inner.ToArray();
    }

}
