using System.Data;
using System.Globalization;
using System.Text;
using DataToExcel.Models;
using DataToExcel.Services.Interfaces;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace DataToExcel.Services;

public class ExcelExportService : IExcelExportService
{
    private readonly IExcelStyleProvider _styleProvider;
    public ExcelExportService(IExcelStyleProvider styleProvider)
        => _styleProvider = styleProvider;

    public Task<ServiceResponse<Stream>> ExportAsync(IEnumerable<IDataRecord> data,
        IReadOnlyList<ColumnDefinition> columns,
        Stream output,
        ExcelExportOptions options,
        CancellationToken ct = default)
        => options.SplitIntoMultipleSheets
            ? ExportMultipleSheetsAsync(data, columns, output, options, ct)
            : ExportAsyncCore(output, options, (worksheetPart, styleMap)
            => WriteWorksheetAsync(worksheetPart, columns, options, styleMap,
                writer =>
                {
                    WriteRows(writer, data, columns, styleMap, ct, ExcelExportLimits.MaxDataRowsPerSheet);
                    return Task.CompletedTask;
                }), ct);

    public Task<ServiceResponse<Stream>> ExportAsync(IAsyncEnumerable<IDataRecord> data,
        IReadOnlyList<ColumnDefinition> columns,
        Stream output,
        ExcelExportOptions options,
        CancellationToken ct = default)
        => options.SplitIntoMultipleSheets
            ? ExportMultipleSheetsAsync(data, columns, output, options, ct)
            : ExportAsyncCore(output, options, (worksheetPart, styleMap)
            => WriteWorksheetAsync(worksheetPart, columns, options, styleMap,
                writer => WriteRows(writer, data, columns, styleMap, ct, ExcelExportLimits.MaxDataRowsPerSheet)), ct);

    private async Task<ServiceResponse<Stream>> ExportMultipleSheetsAsync(IEnumerable<IDataRecord> data,
        IReadOnlyList<ColumnDefinition> columns,
        Stream output,
        ExcelExportOptions options,
        CancellationToken ct)
    {
        try
        {
            if (!output.CanSeek)
                throw new ArgumentException("Stream must be seekable", nameof(output));

            var styleResponse = _styleProvider.BuildStylesheet(out var styleMap);
            if (!styleResponse.IsSuccess || styleResponse.Data is null)
                return new ServiceResponse<Stream> { IsSuccess = false, ErrorMessage = styleResponse.ErrorMessage };
            var stylesheet = styleResponse.Data;

            using var document = SpreadsheetDocument.Create(output, SpreadsheetDocumentType.Workbook, true);
            var workbookPart = document.AddWorkbookPart();
            workbookPart.Workbook = new Workbook();
            var stylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();
            stylesPart.Stylesheet = stylesheet;
            var sheets = workbookPart.Workbook.AppendChild(new Sheets());

            using var enumerator = data.GetEnumerator();
            var bufferedEnumerator = new BufferedRecordEnumerator(enumerator);
            var sheetIndex = 1;
            var hasMore = true;

            do
            {
                var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                await WriteWorksheetAsync(worksheetPart, columns, options, styleMap, writer =>
                {
                    WriteRows(writer, bufferedEnumerator, columns, styleMap, ct, ExcelExportLimits.MaxDataRowsPerSheet);
                    return Task.CompletedTask;
                });

                sheets.AppendChild(new Sheet
                {
                    Id = workbookPart.GetIdOfPart(worksheetPart),
                    SheetId = (uint)sheetIndex,
                    Name = ComposeSheetName(options.SheetName, sheetIndex)
                });

                sheetIndex++;
                hasMore = bufferedEnumerator.TryPeekNext(out _);
            } while (hasMore);

            workbookPart.Workbook.Save();
            await output.FlushAsync(ct);
            return new ServiceResponse<Stream>(output) { IsSuccess = true };
        }
        catch (Exception ex)
        {
            return new ServiceResponse<Stream> { IsSuccess = false, ErrorMessage = ex.Message };
        }
    }

    private async Task<ServiceResponse<Stream>> ExportMultipleSheetsAsync(IAsyncEnumerable<IDataRecord> data,
        IReadOnlyList<ColumnDefinition> columns,
        Stream output,
        ExcelExportOptions options,
        CancellationToken ct)
    {
        try
        {
            if (!output.CanSeek)
                throw new ArgumentException("Stream must be seekable", nameof(output));

            var styleResponse = _styleProvider.BuildStylesheet(out var styleMap);
            if (!styleResponse.IsSuccess || styleResponse.Data is null)
                return new ServiceResponse<Stream> { IsSuccess = false, ErrorMessage = styleResponse.ErrorMessage };
            var stylesheet = styleResponse.Data;

            using var document = SpreadsheetDocument.Create(output, SpreadsheetDocumentType.Workbook, true);
            var workbookPart = document.AddWorkbookPart();
            workbookPart.Workbook = new Workbook();
            var stylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();
            stylesPart.Stylesheet = stylesheet;
            var sheets = workbookPart.Workbook.AppendChild(new Sheets());

            await using var enumerator = data.GetAsyncEnumerator(ct);
            var bufferedEnumerator = new BufferedAsyncRecordEnumerator(enumerator);
            var sheetIndex = 1;
            var hasMore = true;

            do
            {
                var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                await WriteWorksheetAsync(worksheetPart, columns, options, styleMap, async writer =>
                {
                    await WriteRows(writer, bufferedEnumerator, columns, styleMap, ct, ExcelExportLimits.MaxDataRowsPerSheet);
                });

                sheets.AppendChild(new Sheet
                {
                    Id = workbookPart.GetIdOfPart(worksheetPart),
                    SheetId = (uint)sheetIndex,
                    Name = ComposeSheetName(options.SheetName, sheetIndex)
                });

                sheetIndex++;
                hasMore = await bufferedEnumerator.TryPeekNextAsync();
            } while (hasMore);

            workbookPart.Workbook.Save();
            await output.FlushAsync(ct);
            return new ServiceResponse<Stream>(output) { IsSuccess = true };
        }
        catch (Exception ex)
        {
            return new ServiceResponse<Stream> { IsSuccess = false, ErrorMessage = ex.Message };
        }
    }

    private async Task<ServiceResponse<Stream>> ExportAsyncCore(Stream output,
        ExcelExportOptions options,
        Func<WorksheetPart, IReadOnlyDictionary<PredefinedStyle, uint>, Task> writeWorksheetAsync,
        CancellationToken ct)
    {
        try
        {
            if (!output.CanSeek)
                throw new ArgumentException("Stream must be seekable", nameof(output));

            var styleResponse = _styleProvider.BuildStylesheet(out var styleMap);
            if (!styleResponse.IsSuccess || styleResponse.Data is null)
                return new ServiceResponse<Stream> { IsSuccess = false, ErrorMessage = styleResponse.ErrorMessage };
            var stylesheet = styleResponse.Data;

            using var document = SpreadsheetDocument.Create(output, SpreadsheetDocumentType.Workbook, true);
            var workbookPart = document.AddWorkbookPart();
            workbookPart.Workbook = new Workbook();
            var stylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();
            stylesPart.Stylesheet = stylesheet;
            var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();

            await writeWorksheetAsync(worksheetPart, styleMap);

            var sheets = workbookPart.Workbook.AppendChild(new Sheets());
            sheets.AppendChild(new Sheet
            {
                Id = workbookPart.GetIdOfPart(worksheetPart),
                SheetId = 1,
                Name = options.SheetName
            });
            workbookPart.Workbook.Save();
            await output.FlushAsync(ct);
            return new ServiceResponse<Stream>(output) { IsSuccess = true };
        }
        catch (Exception ex)
        {
            return new ServiceResponse<Stream> { IsSuccess = false, ErrorMessage = ex.Message };
        }
    }

    private static async Task WriteWorksheetAsync(WorksheetPart worksheetPart,
        IReadOnlyList<ColumnDefinition> columns,
        ExcelExportOptions options,
        IReadOnlyDictionary<PredefinedStyle, uint> styleMap,
        Func<OpenXmlWriter, Task> writeRowsAsync)
    {
        using var writer = OpenXmlWriter.Create(worksheetPart);
        writer.WriteStartElement(new Worksheet());

        WriteSheetViews(writer, options);
        WriteColumns(writer, columns);
        WriteSheetFormatProperties(writer, columns);

        writer.WriteStartElement(new SheetData());
        WriteHeader(writer, columns, styleMap);
        await writeRowsAsync(writer);
        writer.WriteEndElement(); // SheetData

        WriteAutoFilter(writer, options, columns.Count);

        writer.WriteEndElement(); // Worksheet
        writer.Close();
    }

    private static void WriteSheetViews(OpenXmlWriter writer, ExcelExportOptions options)
    {
        if (!options.FreezeHeader) return;
        writer.WriteStartElement(new SheetViews());
        writer.WriteElement(new SheetView
        {
            WorkbookViewId = 0,
            Pane = new Pane
            {
                VerticalSplit = 1,
                TopLeftCell = "A2",
                ActivePane = PaneValues.BottomLeft,
                State = PaneStateValues.Frozen
            }
        });
        writer.WriteEndElement(); // SheetViews
    }

    private static void WriteSheetFormatProperties(OpenXmlWriter writer, IReadOnlyList<ColumnDefinition> columns)
    {
        if (!columns.Any(c => c.Group)) return;
        writer.WriteElement(new SheetFormatProperties { OutlineLevelRow = 1 });
    }

    private static void WriteColumns(OpenXmlWriter writer, IReadOnlyList<ColumnDefinition> columns)
    {
        if (!columns.Any(c => c.Width.HasValue || c.Hidden)) return;
        writer.WriteStartElement(new Columns());
        uint i = 1;
        foreach (var col in columns)
        {
            if (col.Width.HasValue || col.Hidden)
            {
                var column = new Column
                {
                    Min = i,
                    Max = i
                };
                if (col.Hidden)
                {
                    column.Hidden = true;
                }
                if (col.Width.HasValue)
                {
                    column.Width = col.Width.Value;
                    column.CustomWidth = true;
                }
                writer.WriteElement(column);
            }
            i++;
        }
        writer.WriteEndElement(); // Columns
    }

    private static void WriteHeader(OpenXmlWriter writer,
        IReadOnlyList<ColumnDefinition> columns,
        IReadOnlyDictionary<PredefinedStyle, uint> styleMap)
    {
        writer.WriteStartElement(new Row());
        foreach (var col in columns)
        {
            writer.WriteElement(new Cell
            {
                DataType = CellValues.String,
                CellValue = new CellValue(col.Title ?? string.Empty),
                StyleIndex = styleMap[PredefinedStyle.Header]
            });
        }
        writer.WriteEndElement(); // Row
    }

    private static void WriteRows(OpenXmlWriter writer,
        IEnumerable<IDataRecord> data,
        IReadOnlyList<ColumnDefinition> columns,
        IReadOnlyDictionary<PredefinedStyle, uint> styleMap,
        CancellationToken ct,
        int maxRows)
    {
        var (groupIndexValue, groupField) = GetGroupInfo(columns);
        object? currentGroup = null;
        var written = 0;

        foreach (var record in data)
        {
            ct.ThrowIfCancellationRequested();
            if (written >= maxRows)
                throw new InvalidOperationException($"Row limit exceeded ({ExcelExportLimits.MaxRowsPerSheet}). Enable splitting to export more rows.");

            var dataRow = CreateDisconnectedRow(record, columns);
            var isGroupRow = IsNewGroupRow(dataRow, groupField, currentGroup, out var newGroupValue);
            if (isGroupRow)
                currentGroup = newGroupValue;

            var row = CreateRow(groupField is not null, isGroupRow);
            writer.WriteStartElement(row);
            WriteRowCells(writer, dataRow, columns, styleMap, groupField, groupIndexValue, isGroupRow);
            writer.WriteEndElement();
            ClearDataRow(dataRow);
            written++;
        }
    }

    private static void WriteRows(OpenXmlWriter writer,
        BufferedRecordEnumerator data,
        IReadOnlyList<ColumnDefinition> columns,
        IReadOnlyDictionary<PredefinedStyle, uint> styleMap,
        CancellationToken ct,
        int maxRows)
    {
        var (groupIndexValue, groupField) = GetGroupInfo(columns);
        object? currentGroup = null;
        var written = 0;

        while (written < maxRows && data.TryGetNext(out var record))
        {
            ct.ThrowIfCancellationRequested();

            var dataRow = CreateDisconnectedRow(record, columns);
            var isGroupRow = IsNewGroupRow(dataRow, groupField, currentGroup, out var newGroupValue);
            if (isGroupRow)
                currentGroup = newGroupValue;

            var row = CreateRow(groupField is not null, isGroupRow);
            writer.WriteStartElement(row);
            WriteRowCells(writer, dataRow, columns, styleMap, groupField, groupIndexValue, isGroupRow);
            writer.WriteEndElement();
            ClearDataRow(dataRow);
            written++;
        }
    }

    private static async Task WriteRows(OpenXmlWriter writer,
        IAsyncEnumerable<IDataRecord> data,
        IReadOnlyList<ColumnDefinition> columns,
        IReadOnlyDictionary<PredefinedStyle, uint> styleMap,
        CancellationToken ct,
        int maxRows)
    {
        var (groupIndexValue, groupField) = GetGroupInfo(columns);
        object? currentGroup = null;
        var written = 0;

        await foreach (var record in data.WithCancellation(ct))
        {
            if (written >= maxRows)
                throw new InvalidOperationException($"Row limit exceeded ({ExcelExportLimits.MaxRowsPerSheet}). Enable splitting to export more rows.");

            var dataRow = CreateDisconnectedRow(record, columns);
            var isGroupRow = IsNewGroupRow(dataRow, groupField, currentGroup, out var newGroupValue);
            if (isGroupRow)
                currentGroup = newGroupValue;

            var row = CreateRow(groupField is not null, isGroupRow);
            writer.WriteStartElement(row);
            WriteRowCells(writer, dataRow, columns, styleMap, groupField, groupIndexValue, isGroupRow);
            writer.WriteEndElement();
            ClearDataRow(dataRow);
            written++;
        }
    }

    private static async Task WriteRows(OpenXmlWriter writer,
        BufferedAsyncRecordEnumerator data,
        IReadOnlyList<ColumnDefinition> columns,
        IReadOnlyDictionary<PredefinedStyle, uint> styleMap,
        CancellationToken ct,
        int maxRows)
    {
        var (groupIndexValue, groupField) = GetGroupInfo(columns);
        object? currentGroup = null;
        var written = 0;

        while (written < maxRows && await data.TryGetNextAsync())
        {
            ct.ThrowIfCancellationRequested();
            var record = data.Current ?? throw new InvalidOperationException("Expected record instance.");

            var dataRow = CreateDisconnectedRow(record, columns);
            var isGroupRow = IsNewGroupRow(dataRow, groupField, currentGroup, out var newGroupValue);
            if (isGroupRow)
                currentGroup = newGroupValue;

            var row = CreateRow(groupField is not null, isGroupRow);
            writer.WriteStartElement(row);
            WriteRowCells(writer, dataRow, columns, styleMap, groupField, groupIndexValue, isGroupRow);
            writer.WriteEndElement();
            ClearDataRow(dataRow);
            written++;
        }
    }

    private static string ComposeSheetName(string sheetName, int sheetIndex)
    {
        var cleanedName = string.IsNullOrWhiteSpace(sheetName) ? "Sheet" : sheetName.Trim();
        if (sheetIndex == 1)
            return TrimSheetName(cleanedName, 0);

        var suffix = $" ({sheetIndex})";
        return $"{TrimSheetName(cleanedName, suffix.Length)}{suffix}";
    }

    private static string TrimSheetName(string sheetName, int suffixLength)
    {
        const int maxLength = 31;
        var available = Math.Max(1, maxLength - suffixLength);
        if (sheetName.Length <= available)
            return sheetName;
        return sheetName[..available];
    }

    private sealed class BufferedRecordEnumerator
    {
        private readonly IEnumerator<IDataRecord> _inner;
        private bool _hasBuffered;
        private IDataRecord? _buffered;

        public BufferedRecordEnumerator(IEnumerator<IDataRecord> inner)
            => _inner = inner;

        public bool TryGetNext(out IDataRecord record)
        {
            if (_hasBuffered)
            {
                record = _buffered ?? throw new InvalidOperationException("Buffered record expected.");
                _buffered = null;
                _hasBuffered = false;
                return true;
            }

            if (_inner.MoveNext())
            {
                record = _inner.Current;
                return true;
            }

            record = null!;
            return false;
        }

        public bool TryPeekNext(out IDataRecord? record)
        {
            if (_hasBuffered)
            {
                record = _buffered;
                return true;
            }

            if (_inner.MoveNext())
            {
                _buffered = _inner.Current;
                _hasBuffered = true;
                record = _buffered;
                return true;
            }

            record = null;
            return false;
        }
    }

    private sealed class BufferedAsyncRecordEnumerator
    {
        private readonly IAsyncEnumerator<IDataRecord> _inner;
        private bool _hasBuffered;
        public IDataRecord? Current { get; private set; }

        public BufferedAsyncRecordEnumerator(IAsyncEnumerator<IDataRecord> inner)
            => _inner = inner;

        public async Task<bool> TryGetNextAsync()
        {
            if (_hasBuffered)
            {
                _hasBuffered = false;
                return true;
            }

            if (await _inner.MoveNextAsync())
            {
                Current = _inner.Current;
                return true;
            }

            Current = null;
            return false;
        }

        public async Task<bool> TryPeekNextAsync()
        {
            if (_hasBuffered)
                return true;

            if (await _inner.MoveNextAsync())
            {
                Current = _inner.Current;
                _hasBuffered = true;
                return true;
            }

            Current = null;
            return false;
        }
    }

    private static (int groupIndex, string? groupField) GetGroupInfo(IReadOnlyList<ColumnDefinition> columns)
    {
        var groupInfo = columns.Select((c, i) => new { c, i }).FirstOrDefault(x => x.c.Group);
        if (groupInfo is null)
            return (-1, null);
        return (groupInfo.i, columns[groupInfo.i].FieldName);
    }

    private static bool IsNewGroupRow(DataRow dataRow, string? groupField, object? currentGroup, out object? newGroupValue)
    {
        newGroupValue = currentGroup;
        if (groupField is null)
            return false;

        var value = dataRow[groupField];
        if (Equals(value, currentGroup))
            return false;

        newGroupValue = value;
        return true;
    }

    private static Row CreateRow(bool hasGroup, bool isGroupRow)
    {
        var row = new Row();
        if (hasGroup && !isGroupRow)
            row.OutlineLevel = 1;
        return row;
    }

    private static void WriteRowCells(OpenXmlWriter writer,
        DataRow dataRow,
        IReadOnlyList<ColumnDefinition> columns,
        IReadOnlyDictionary<PredefinedStyle, uint> styleMap,
        string? groupField,
        int groupIndexValue,
        bool isGroupRow)
    {
        for (int i = 0; i < columns.Count; i++)
        {
            var col = columns[i];
            if (groupField is not null && i == groupIndexValue && !isGroupRow)
            {
                writer.WriteElement(new Cell());
                continue;
            }

            var value = dataRow[col.FieldName];
            if (value == DBNull.Value || value is null)
            {
                writer.WriteElement(new Cell());
                continue;
            }

            var cell = CreateCell(value, col, styleMap);
            writer.WriteElement(cell);
        }
    }

    private static DataRow CreateDisconnectedRow(IDataRecord record, IReadOnlyList<ColumnDefinition> columns)
    {
        var recordValues = new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase);
        for (int i = 0; i < record.FieldCount; i++)
        {
            recordValues[record.GetName(i)] = record.IsDBNull(i) ? DBNull.Value : record.GetValue(i);
        }

        var table = new DataTable();
        foreach (var col in columns)
        {
            table.Columns.Add(col.FieldName, typeof(object));
        }

        var row = table.NewRow();
        foreach (var col in columns)
        {
            if (recordValues.TryGetValue(col.FieldName, out var value))
            {
                row[col.FieldName] = value ?? DBNull.Value;
            }
            else
            {
                row[col.FieldName] = DBNull.Value;
            }
        }

        return row;
    }

    private static void ClearDataRow(DataRow row)
    {
        row.Table?.Clear();
    }

    private static void WriteAutoFilter(OpenXmlWriter writer, ExcelExportOptions options, int columnCount)
    {
        if (!options.AutoFilter) return;
        var endCol = GetColumnName(columnCount);
        writer.WriteElement(new AutoFilter { Reference = $"A1:{endCol}1" });
    }

    private static Cell CreateCell(object value, ColumnDefinition col,
        IReadOnlyDictionary<PredefinedStyle, uint> styleMap)
    {
        var style = col.Style ?? GetStyleFromDataType(col.DataType);
        var cell = new Cell { StyleIndex = styleMap[style] };
        switch (col.DataType)
        {
            case ColumnDataType.Number:
            case ColumnDataType.Currency:
            case ColumnDataType.Percentage:
                cell.DataType = CellValues.Number;
                cell.CellValue = new CellValue(Convert.ToString(value, CultureInfo.InvariantCulture) ?? string.Empty);
                break;
            case ColumnDataType.DateTime:
                cell.DataType = CellValues.Number;
                var dt = Convert.ToDateTime(value, CultureInfo.InvariantCulture);
                cell.CellValue = new CellValue(dt.ToOADate().ToString(CultureInfo.InvariantCulture));
                break;
            case ColumnDataType.Boolean:
                cell.DataType = CellValues.Boolean;
                cell.CellValue = new CellValue((bool)value ? "1" : "0");
                break;
            default:
                cell.DataType = CellValues.InlineString;
                var s = Convert.ToString(value, CultureInfo.InvariantCulture) ?? string.Empty;
                var inline = new InlineString();
                inline.AppendChild(new Text(s));
                cell.InlineString = inline;
                break;
        }
        return cell;
    }

    private static PredefinedStyle GetStyleFromDataType(ColumnDataType type) => type switch
    {
        ColumnDataType.Number => PredefinedStyle.Number,
        ColumnDataType.DateTime => PredefinedStyle.DateTime,
        ColumnDataType.Boolean => PredefinedStyle.Boolean,
        ColumnDataType.Currency => PredefinedStyle.Currency,
        ColumnDataType.Percentage => PredefinedStyle.Percentage,
        _ => PredefinedStyle.Text
    };

    private static string GetColumnName(int index)
    {
        var dividend = index;
        var sb = new StringBuilder();
        while (dividend > 0)
        {
            var modulo = (dividend - 1) % 26;
            sb.Insert(0, Convert.ToChar(65 + modulo));
            dividend = (dividend - modulo) / 26;
        }
        return sb.ToString();
    }
}
