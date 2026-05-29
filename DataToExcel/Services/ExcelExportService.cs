using System.Data;
using System.Globalization;
using System.IO.Compression;
using System.Text;
using System.Xml;
using DataToExcel.Models;
using DataToExcel.Services.Interfaces;
using DataToExcel.Utilities;
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
        => ExportAsync(AsyncEnumerableHelpers.ToAsyncEnumerable(data, ct), columns, output, options, ct);

    public Task<ServiceResponse<Stream>> ExportAsync(IAsyncEnumerable<IDataRecord> data,
        IReadOnlyList<ColumnDefinition> columns,
        Stream output,
        ExcelExportOptions options,
        CancellationToken ct = default)
        => options.SplitIntoMultipleSheets
            ? ExportMultipleSheetsAsync(data, columns, output, options, ct)
            : ExportAsyncCore(output, options, (worksheetStream, styleMap)
            => WriteWorksheetXmlAsync(worksheetStream, columns, options, styleMap,
                writer => WriteRows(writer, data, columns, styleMap, ExcelExportLimits.MaxDataRowsPerSheet, ct)), ct);

    private Task<ServiceResponse<Stream>> ExportMultipleSheetsAsync(IAsyncEnumerable<IDataRecord> data,
        IReadOnlyList<ColumnDefinition> columns,
        Stream output,
        ExcelExportOptions options,
        CancellationToken ct)
        => ExportMultipleSheetsAsyncCore(output, async (workbookPart, sheets, styleMap) =>
        {
            await using var enumerator = data.GetAsyncEnumerator(ct);
            var bufferedEnumerator = new BufferedAsyncRecordEnumerator(enumerator);
            var sheetIndex = 1;
            var hasMore = true;

            do
            {
                await AddSheetAsync(workbookPart, sheets, sheetIndex, columns, options, styleMap, async writer =>
                {
                    await WriteRows(writer, bufferedEnumerator, columns, styleMap, ExcelExportLimits.MaxDataRowsPerSheet, ct);
                });
                sheetIndex++;
                hasMore = await bufferedEnumerator.TryPeekNextAsync();
            } while (hasMore);
        }, ct);


    private static async Task<ServiceResponse<Stream>> ExportUsingSeekableStreamAsync(
        Stream output,
        Func<Stream, Task> exportToSeekableStreamAsync,
        CancellationToken ct)
    {
        var needsStaging = !output.CanSeek;
        var seekableStream = needsStaging ? CreateTempFileStream() : output;

        try
        {
            await exportToSeekableStreamAsync(seekableStream);

            if (needsStaging)
            {
                seekableStream.Position = 0;
                await seekableStream.CopyToAsync(output, 81920, ct);
                await output.FlushAsync(ct);
            }
            else
            {
                await output.FlushAsync(ct);
            }

            return new ServiceResponse<Stream>(output) { IsSuccess = true };
        }
        finally
        {
            if (needsStaging)
            {
                await seekableStream.DisposeAsync();
            }
        }
    }

    private static FileStream CreateTempFileStream()
    {
        var tempFile = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName());
        return new FileStream(tempFile, FileMode.Create, FileAccess.ReadWrite, FileShare.None, 4096,
            FileOptions.Asynchronous | FileOptions.SequentialScan | FileOptions.DeleteOnClose);
    }

    private async Task<ServiceResponse<Stream>> ExportMultipleSheetsAsyncCore(
        Stream output,
        Func<WorkbookPart, Sheets, IReadOnlyDictionary<PredefinedStyle, uint>, Task> writeSheetsAsync,
        CancellationToken ct)
    {
        try
        {
            var styleResponse = _styleProvider.BuildStylesheet(out var styleMap);
            if (!styleResponse.IsSuccess || styleResponse.Data is null)
                return new ServiceResponse<Stream> { IsSuccess = false, ErrorMessage = styleResponse.ErrorMessage };
            var stylesheet = styleResponse.Data;

            return await ExportUsingSeekableStreamAsync(output, async seekableStream =>
            {
                using var document = SpreadsheetDocument.Create(seekableStream, SpreadsheetDocumentType.Workbook, true);
                var workbookPart = document.AddWorkbookPart();
                workbookPart.Workbook = new Workbook();
                var stylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();
                stylesPart.Stylesheet = stylesheet;
                var sheets = workbookPart.Workbook.AppendChild(new Sheets());

                await writeSheetsAsync(workbookPart, sheets, styleMap);
                workbookPart.Workbook.Save();
            }, ct);
        }
        catch (Exception ex)
        {
            return new ServiceResponse<Stream> { IsSuccess = false, ErrorMessage = ex.Message };
        }
    }

    private async Task<ServiceResponse<Stream>> ExportAsyncCore(Stream output,
        ExcelExportOptions options,
        Func<Stream, IReadOnlyDictionary<PredefinedStyle, uint>, Task> writeWorksheetAsync,
        CancellationToken ct)
    {
        try
        {
            var styleResponse = _styleProvider.BuildStylesheet(out var styleMap);
            if (!styleResponse.IsSuccess || styleResponse.Data is null)
                return new ServiceResponse<Stream> { IsSuccess = false, ErrorMessage = styleResponse.ErrorMessage };
            var stylesheet = styleResponse.Data;

            await ExportSingleSheetPackageAsync(output, options, stylesheet, styleMap, writeWorksheetAsync, ct);
            return new ServiceResponse<Stream>(output) { IsSuccess = true };
        }
        catch (Exception ex)
        {
            return new ServiceResponse<Stream> { IsSuccess = false, ErrorMessage = ex.Message };
        }
    }

    private static async Task ExportSingleSheetPackageAsync(Stream output,
        ExcelExportOptions options,
        Stylesheet stylesheet,
        IReadOnlyDictionary<PredefinedStyle, uint> styleMap,
        Func<Stream, IReadOnlyDictionary<PredefinedStyle, uint>, Task> writeWorksheetAsync,
        CancellationToken ct)
    {
        using var archive = new ZipArchive(output, ZipArchiveMode.Create, leaveOpen: true);

        await WriteContentTypesAsync(archive, ct);
        await WriteRootRelationshipsAsync(archive, ct);
        await WriteWorkbookAsync(archive, ComposeSheetName(options.SheetName, 1), ct);
        await WriteWorkbookRelationshipsAsync(archive, ct);
        await WriteStylesAsync(archive, stylesheet, ct);

        var worksheetEntry = archive.CreateEntry("xl/worksheets/sheet1.xml", CompressionLevel.Fastest);
        await using var worksheetStream = worksheetEntry.Open();
        await writeWorksheetAsync(worksheetStream, styleMap);

        await output.FlushAsync(ct);
    }

    private static async Task WriteContentTypesAsync(ZipArchive archive, CancellationToken ct)
    {
        var entry = archive.CreateEntry("[Content_Types].xml", CompressionLevel.Fastest);
        await using var stream = entry.Open();
        await using var writer = CreatePackageXmlWriter(stream);

        await writer.WriteStartDocumentAsync();
        await writer.WriteStartElementAsync(null, "Types", "http://schemas.openxmlformats.org/package/2006/content-types");
        await WriteDefaultContentTypeAsync(writer, "rels", "application/vnd.openxmlformats-package.relationships+xml");
        await WriteDefaultContentTypeAsync(writer, "xml", "application/xml");
        await WriteOverrideContentTypeAsync(writer, "/xl/workbook.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml");
        await WriteOverrideContentTypeAsync(writer, "/xl/worksheets/sheet1.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml");
        await WriteOverrideContentTypeAsync(writer, "/xl/styles.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml");
        await writer.WriteEndElementAsync();
        await writer.WriteEndDocumentAsync();
        await writer.FlushAsync();
        ct.ThrowIfCancellationRequested();
    }

    private static async Task WriteRootRelationshipsAsync(ZipArchive archive, CancellationToken ct)
    {
        var entry = archive.CreateEntry("_rels/.rels", CompressionLevel.Fastest);
        await using var stream = entry.Open();
        await using var writer = CreatePackageXmlWriter(stream);

        await writer.WriteStartDocumentAsync();
        await writer.WriteStartElementAsync(null, "Relationships", "http://schemas.openxmlformats.org/package/2006/relationships");
        await WriteRelationshipAsync(writer, "rId1",
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument",
            "xl/workbook.xml");
        await writer.WriteEndElementAsync();
        await writer.WriteEndDocumentAsync();
        await writer.FlushAsync();
        ct.ThrowIfCancellationRequested();
    }

    private static async Task WriteWorkbookAsync(ZipArchive archive, string sheetName, CancellationToken ct)
    {
        var entry = archive.CreateEntry("xl/workbook.xml", CompressionLevel.Fastest);
        await using var stream = entry.Open();
        await using var writer = CreatePackageXmlWriter(stream);

        await writer.WriteStartDocumentAsync();
        await writer.WriteStartElementAsync(null, "workbook", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
        await writer.WriteAttributeStringAsync("xmlns", "r", null, "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
        await writer.WriteStartElementAsync(null, "sheets", null);
        await writer.WriteStartElementAsync(null, "sheet", null);
        await writer.WriteAttributeStringAsync(null, "name", null, sheetName);
        await writer.WriteAttributeStringAsync(null, "sheetId", null, "1");
        await writer.WriteAttributeStringAsync("r", "id", null, "rId1");
        await writer.WriteEndElementAsync(); // sheet
        await writer.WriteEndElementAsync(); // sheets
        await writer.WriteEndElementAsync(); // workbook
        await writer.WriteEndDocumentAsync();
        await writer.FlushAsync();
        ct.ThrowIfCancellationRequested();
    }

    private static async Task WriteWorkbookRelationshipsAsync(ZipArchive archive, CancellationToken ct)
    {
        var entry = archive.CreateEntry("xl/_rels/workbook.xml.rels", CompressionLevel.Fastest);
        await using var stream = entry.Open();
        await using var writer = CreatePackageXmlWriter(stream);

        await writer.WriteStartDocumentAsync();
        await writer.WriteStartElementAsync(null, "Relationships", "http://schemas.openxmlformats.org/package/2006/relationships");
        await WriteRelationshipAsync(writer, "rId1",
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet",
            "worksheets/sheet1.xml");
        await WriteRelationshipAsync(writer, "rId2",
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles",
            "styles.xml");
        await writer.WriteEndElementAsync();
        await writer.WriteEndDocumentAsync();
        await writer.FlushAsync();
        ct.ThrowIfCancellationRequested();
    }

    private static async Task WriteStylesAsync(ZipArchive archive, Stylesheet stylesheet, CancellationToken ct)
    {
        var entry = archive.CreateEntry("xl/styles.xml", CompressionLevel.Fastest);
        await using var stream = entry.Open();
        stylesheet.Save(stream);
        ct.ThrowIfCancellationRequested();
    }

    private static XmlWriter CreatePackageXmlWriter(Stream stream)
        => XmlWriter.Create(stream, new XmlWriterSettings
        {
            Async = true,
            Encoding = Encoding.UTF8,
            CloseOutput = false
        });

    private static async Task WriteDefaultContentTypeAsync(XmlWriter writer, string extension, string contentType)
    {
        await writer.WriteStartElementAsync(null, "Default", null);
        await writer.WriteAttributeStringAsync(null, "Extension", null, extension);
        await writer.WriteAttributeStringAsync(null, "ContentType", null, contentType);
        await writer.WriteEndElementAsync();
    }

    private static async Task WriteOverrideContentTypeAsync(XmlWriter writer, string partName, string contentType)
    {
        await writer.WriteStartElementAsync(null, "Override", null);
        await writer.WriteAttributeStringAsync(null, "PartName", null, partName);
        await writer.WriteAttributeStringAsync(null, "ContentType", null, contentType);
        await writer.WriteEndElementAsync();
    }

    private static async Task WriteRelationshipAsync(XmlWriter writer, string id, string type, string target)
    {
        await writer.WriteStartElementAsync(null, "Relationship", null);
        await writer.WriteAttributeStringAsync(null, "Id", null, id);
        await writer.WriteAttributeStringAsync(null, "Type", null, type);
        await writer.WriteAttributeStringAsync(null, "Target", null, target);
        await writer.WriteEndElementAsync();
    }

    private static async Task WriteWorksheetAsync(WorksheetPart worksheetPart,
        IReadOnlyList<ColumnDefinition> columns,
        ExcelExportOptions options,
        IReadOnlyDictionary<PredefinedStyle, uint> styleMap,
        Func<XmlWriter, Task> writeRowsAsync)
    {
        await using var stream = worksheetPart.GetStream(FileMode.Create, FileAccess.Write);
        await WriteWorksheetXmlAsync(stream, columns, options, styleMap, writeRowsAsync);
    }

    private static async Task WriteWorksheetXmlAsync(Stream stream,
        IReadOnlyList<ColumnDefinition> columns,
        ExcelExportOptions options,
        IReadOnlyDictionary<PredefinedStyle, uint> styleMap,
        Func<XmlWriter, Task> writeRowsAsync)
    {
        await using var writer = XmlWriter.Create(stream, new XmlWriterSettings
        {
            Async = true,
            Encoding = Encoding.UTF8,
            CloseOutput = false
        });
        await writer.WriteStartDocumentAsync();
        await writer.WriteStartElementAsync(null, "worksheet", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");

        WriteSheetViews(writer, options);
        WriteColumns(writer, columns);
        WriteSheetFormatProperties(writer, columns);

        writer.WriteStartElement("sheetData");
        WriteHeader(writer, columns, styleMap);
        await writeRowsAsync(writer);
        writer.WriteEndElement(); // sheetData

        WriteAutoFilter(writer, options, columns.Count);

        writer.WriteEndElement(); // worksheet
        await writer.WriteEndDocumentAsync();
        await writer.FlushAsync();
    }

    private static void WriteSheetViews(XmlWriter writer, ExcelExportOptions options)
    {
        if (!options.FreezeHeader) return;
        writer.WriteStartElement("sheetViews");
        writer.WriteStartElement("sheetView");
        writer.WriteAttributeString("workbookViewId", "0");
        writer.WriteStartElement("pane");
        writer.WriteAttributeString("ySplit", "1");
        writer.WriteAttributeString("topLeftCell", "A2");
        writer.WriteAttributeString("activePane", "bottomLeft");
        writer.WriteAttributeString("state", "frozen");
        writer.WriteEndElement(); // pane
        writer.WriteEndElement(); // sheetView
        writer.WriteEndElement(); // sheetViews
    }

    private static void WriteSheetFormatProperties(XmlWriter writer, IReadOnlyList<ColumnDefinition> columns)
    {
        if (!columns.Any(c => c.Group)) return;
        writer.WriteStartElement("sheetFormatPr");
        writer.WriteAttributeString("outlineLevelRow", "1");
        writer.WriteEndElement();
    }

    private static void WriteColumns(XmlWriter writer, IReadOnlyList<ColumnDefinition> columns)
    {
        if (!columns.Any(c => c.Width.HasValue || c.Hidden)) return;
        writer.WriteStartElement("cols");
        uint i = 1;
        foreach (var col in columns)
        {
            if (col.Width.HasValue || col.Hidden)
            {
                writer.WriteStartElement("col");
                writer.WriteAttributeString("min", i.ToString(CultureInfo.InvariantCulture));
                writer.WriteAttributeString("max", i.ToString(CultureInfo.InvariantCulture));
                if (col.Hidden)
                {
                    writer.WriteAttributeString("hidden", "1");
                }
                if (col.Width.HasValue)
                {
                    writer.WriteAttributeString("width", col.Width.Value.ToString(CultureInfo.InvariantCulture));
                    writer.WriteAttributeString("customWidth", "1");
                }
                writer.WriteEndElement();
            }
            i++;
        }
        writer.WriteEndElement(); // cols
    }

    private static void WriteHeader(XmlWriter writer,
        IReadOnlyList<ColumnDefinition> columns,
        IReadOnlyDictionary<PredefinedStyle, uint> styleMap)
    {
        writer.WriteStartElement("row");
        foreach (var col in columns)
        {
            WriteCell(writer, col.Title ?? string.Empty, CellValues.String, styleMap[PredefinedStyle.Header]);
        }
        writer.WriteEndElement(); // row
    }

    private static async Task WriteRows(XmlWriter writer,
        IAsyncEnumerable<IDataRecord> data,
        IReadOnlyList<ColumnDefinition> columns,
        IReadOnlyDictionary<PredefinedStyle, uint> styleMap,
        int maxRows,
        CancellationToken ct)
    {
        await using var enumerator = data.GetAsyncEnumerator(ct);
        var context = new WriteRowsContext(columns, styleMap, maxRows, EnforceLimit: true, ct);
        await WriteRowsCoreAsync(writer, context,
            moveNextAsync: () => enumerator.MoveNextAsync().AsTask(),
            current: () => enumerator.Current);
    }

    private static async Task WriteRows(XmlWriter writer,
        BufferedAsyncRecordEnumerator data,
        IReadOnlyList<ColumnDefinition> columns,
        IReadOnlyDictionary<PredefinedStyle, uint> styleMap,
        int maxRows,
        CancellationToken ct)
    {
        var context = new WriteRowsContext(columns, styleMap, maxRows, EnforceLimit: false, ct);
        await WriteRowsCoreAsync(writer, context,
            moveNextAsync: data.TryGetNextAsync,
            current: () => data.Current);
    }

    private static async Task AddSheetAsync(WorkbookPart workbookPart,
        Sheets sheets,
        int sheetIndex,
        IReadOnlyList<ColumnDefinition> columns,
        ExcelExportOptions options,
        IReadOnlyDictionary<PredefinedStyle, uint> styleMap,
        Func<XmlWriter, Task> writeRowsAsync)
    {
        var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
        await WriteWorksheetAsync(worksheetPart, columns, options, styleMap, writeRowsAsync);

        sheets.AppendChild(new Sheet
        {
            Id = workbookPart.GetIdOfPart(worksheetPart),
            SheetId = (uint)sheetIndex,
            Name = ComposeSheetName(options.SheetName, sheetIndex)
        });
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

    private static (int groupIndex, string? groupField) GetGroupInfo(IReadOnlyList<ColumnDefinition> columns)
    {
        var groupInfo = columns.Select((c, i) => new { c, i }).FirstOrDefault(x => x.c.Group);
        if (groupInfo is null)
            return (-1, null);
        return (groupInfo.i, columns[groupInfo.i].FieldName);
    }

    private static async Task WriteRowsCoreAsync(XmlWriter writer,
        WriteRowsContext context,
        Func<Task<bool>> moveNextAsync,
        Func<IDataRecord?> current)
    {
        var (groupIndexValue, groupField) = GetGroupInfo(context.Columns);
        object? currentGroup = null;
        var hasCurrentGroup = false;
        var written = 0;
        int[]? ordinals = null;

        while (await moveNextAsync())
        {
            context.CancellationToken.ThrowIfCancellationRequested();
            if (written >= context.MaxRows)
            {
                if (context.EnforceLimit)
                {
                    throw new InvalidOperationException(
                        $"Row limit exceeded ({ExcelExportLimits.MaxRowsPerSheet}). Enable splitting to export more rows.");
                }
                break;
            }

            var record = current() ?? throw new InvalidOperationException("Expected record instance.");
            ordinals ??= BuildOrdinals(record, context.Columns);
            var isGroupRow = IsNewGroupRow(record, ordinals, groupField, groupIndexValue, currentGroup,
                hasCurrentGroup, out var newGroupValue);
            if (isGroupRow)
            {
                currentGroup = newGroupValue;
                hasCurrentGroup = true;
            }

            WriteRow(writer, record, context.Columns, ordinals, context.StyleMap, groupField, groupIndexValue,
                isGroupRow);
            written++;
        }
    }

    private readonly record struct WriteRowsContext(
        IReadOnlyList<ColumnDefinition> Columns,
        IReadOnlyDictionary<PredefinedStyle, uint> StyleMap,
        int MaxRows,
        bool EnforceLimit,
        CancellationToken CancellationToken);

    private static void WriteRow(XmlWriter writer,
        IDataRecord record,
        IReadOnlyList<ColumnDefinition> columns,
        IReadOnlyList<int> ordinals,
        IReadOnlyDictionary<PredefinedStyle, uint> styleMap,
        string? groupField,
        int groupIndexValue,
        bool isGroupRow)
    {
        WriteRowStart(writer, groupField is not null, isGroupRow);
        WriteRowCells(writer, record, columns, ordinals, styleMap, groupField, groupIndexValue, isGroupRow);
        writer.WriteEndElement();
    }

    private static int[] BuildOrdinals(IDataRecord record, IReadOnlyList<ColumnDefinition> columns)
    {
        var ordinalsByName = BuildOrdinalMap(record);
        var ordinals = new int[columns.Count];
        for (int i = 0; i < columns.Count; i++)
        {
            ordinals[i] = TryGetOrdinal(record, ordinalsByName, columns[i].FieldName);
        }
        return ordinals;
    }

    private static Dictionary<string, int> BuildOrdinalMap(IDataRecord record)
    {
        var map = new Dictionary<string, int>(record.FieldCount, StringComparer.OrdinalIgnoreCase);
        for (var i = 0; i < record.FieldCount; i++)
        {
            map[record.GetName(i)] = i;
        }

        return map;
    }

    private static int TryGetOrdinal(IDataRecord record, IReadOnlyDictionary<string, int> ordinalsByName, string fieldName)
    {
        if (ordinalsByName.TryGetValue(fieldName, out var ordinal))
            return ordinal;

        try
        {
            return record.GetOrdinal(fieldName);
        }
        catch (IndexOutOfRangeException)
        {
            return -1;
        }
    }

    private static object? GetRecordValue(IDataRecord record, int ordinal)
    {
        if (ordinal < 0 || record.IsDBNull(ordinal))
            return null;
        return record.GetValue(ordinal);
    }

    private static bool IsNewGroupRow(IDataRecord record,
        IReadOnlyList<int> ordinals,
        string? groupField,
        int groupIndex,
        object? currentGroup,
        bool hasCurrentGroup,
        out object? newGroupValue)
    {
        newGroupValue = currentGroup;
        if (groupField is null || groupIndex < 0)
            return false;

        var value = GetRecordValue(record, ordinals[groupIndex]);
        if (hasCurrentGroup && Equals(value, currentGroup))
            return false;

        newGroupValue = value;
        return true;
    }

    private static void WriteRowStart(XmlWriter writer, bool hasGroup, bool isGroupRow)
    {
        writer.WriteStartElement("row");
        if (hasGroup && !isGroupRow)
            writer.WriteAttributeString("outlineLevel", "1");
    }

    private static void WriteRowCells(XmlWriter writer,
        IDataRecord record,
        IReadOnlyList<ColumnDefinition> columns,
        IReadOnlyList<int> ordinals,
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
                WriteBlankCell(writer);
                continue;
            }

            var ordinal = ordinals[i];
            if (ordinal < 0 || record.IsDBNull(ordinal))
            {
                WriteBlankCell(writer);
                continue;
            }

            WriteCell(writer, record, ordinal, col, styleMap);
        }
    }

    private static void WriteAutoFilter(XmlWriter writer, ExcelExportOptions options, int columnCount)
    {
        if (!options.AutoFilter) return;
        var endCol = GetColumnName(columnCount);
        writer.WriteStartElement("autoFilter");
        writer.WriteAttributeString("ref", $"A1:{endCol}1");
        writer.WriteEndElement();
    }

    private static void WriteCell(XmlWriter writer, IDataRecord record, int ordinal, ColumnDefinition col,
        IReadOnlyDictionary<PredefinedStyle, uint> styleMap)
    {
        var style = col.Style ?? GetStyleFromDataType(col.DataType);
        switch (col.DataType)
        {
            case ColumnDataType.Number:
            case ColumnDataType.Currency:
            case ColumnDataType.Percentage:
                WriteNumberCell(writer, record, ordinal, styleMap[style]);
                break;
            case ColumnDataType.DateTime:
                WriteDateTimeCell(writer, record, ordinal, styleMap[style]);
                break;
            case ColumnDataType.Boolean:
                WriteCell(writer, GetBooleanCellValue(record, ordinal), CellValues.Boolean, styleMap[style]);
                break;
            default:
                WriteCell(writer, GetStringCellValue(record, ordinal), CellValues.String, styleMap[style]);
                break;
        }
    }

    private static void WriteNumberCell(XmlWriter writer, IDataRecord record, int ordinal, uint styleIndex)
    {
        writer.WriteStartElement("c");
        writer.WriteAttributeString("s", styleIndex.ToString(CultureInfo.InvariantCulture));
        writer.WriteAttributeString("t", "n");
        writer.WriteStartElement("v");

        var fieldType = record.GetFieldType(ordinal);
        if (fieldType == typeof(int))
            writer.WriteValue(record.GetInt32(ordinal));
        else if (fieldType == typeof(long))
            writer.WriteValue(record.GetInt64(ordinal));
        else if (fieldType == typeof(short))
            writer.WriteValue(record.GetInt16(ordinal));
        else if (fieldType == typeof(decimal))
            writer.WriteValue(record.GetDecimal(ordinal));
        else if (fieldType == typeof(double))
            writer.WriteValue(record.GetDouble(ordinal));
        else if (fieldType == typeof(float))
            writer.WriteValue(record.GetFloat(ordinal));
        else
            writer.WriteString(Convert.ToString(record.GetValue(ordinal), CultureInfo.InvariantCulture) ?? string.Empty);

        writer.WriteEndElement(); // v
        writer.WriteEndElement(); // c
    }

    private static void WriteDateTimeCell(XmlWriter writer, IDataRecord record, int ordinal, uint styleIndex)
    {
        var value = record.GetFieldType(ordinal) == typeof(DateTime)
            ? record.GetDateTime(ordinal)
            : Convert.ToDateTime(record.GetValue(ordinal), CultureInfo.InvariantCulture);
        WriteCell(writer, value.ToOADate().ToString(CultureInfo.InvariantCulture), CellValues.Number, styleIndex);
    }

    private static string GetBooleanCellValue(IDataRecord record, int ordinal)
    {
        if (record.GetFieldType(ordinal) == typeof(bool))
            return record.GetBoolean(ordinal) ? "1" : "0";
        return Convert.ToBoolean(record.GetValue(ordinal), CultureInfo.InvariantCulture) ? "1" : "0";
    }

    private static string GetStringCellValue(IDataRecord record, int ordinal)
    {
        if (record.GetFieldType(ordinal) == typeof(string))
            return record.GetString(ordinal);
        return Convert.ToString(record.GetValue(ordinal), CultureInfo.InvariantCulture) ?? string.Empty;
    }

    private static void WriteCell(XmlWriter writer, string value, CellValues dataType, uint styleIndex)
    {
        writer.WriteStartElement("c");
        writer.WriteAttributeString("s", styleIndex.ToString(CultureInfo.InvariantCulture));
        writer.WriteAttributeString("t", GetCellDataTypeValue(dataType));
        writer.WriteStartElement("v");
        writer.WriteString(value);
        writer.WriteEndElement(); // v
        writer.WriteEndElement(); // c
    }

    private static void WriteBlankCell(XmlWriter writer)
    {
        writer.WriteStartElement("c");
        writer.WriteEndElement();
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

    private static string GetCellDataTypeValue(CellValues dataType)
    {
        if (dataType == CellValues.Number)
            return "n";
        if (dataType == CellValues.Boolean)
            return "b";
        return "str";
    }

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
