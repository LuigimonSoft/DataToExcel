using System.Data;
using System.Globalization;
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

    public async Task<ServiceResponse<Stream>> ExportAsync(IEnumerable<IDataRecord> data,
        IReadOnlyList<ColumnDefinition> columns,
        Stream output,
        ExcelExportOptions options,
        CancellationToken ct = default)
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

        using (var writer = OpenXmlWriter.Create(worksheetPart))
        {
            writer.WriteStartElement(new Worksheet());

            if (options.FreezeHeader)
            {
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

            if (columns.Any(c => c.Width.HasValue))
            {
                writer.WriteStartElement(new Columns());
                uint i = 1;
                foreach (var col in columns)
                {
                    if (col.Width.HasValue)
                    {
                        writer.WriteElement(new Column
                        {
                            Min = i,
                            Max = i,
                            Width = col.Width.Value,
                            CustomWidth = true
                        });
                    }
                    i++;
                }
                writer.WriteEndElement(); // Columns
            }

            writer.WriteStartElement(new SheetData());
            // Header
            writer.WriteStartElement(new Row());
            foreach (var col in columns)
            {
                writer.WriteElement(new Cell
                {
                    DataType = CellValues.String,
                    CellValue = new CellValue(col.Title),
                    StyleIndex = styleMap[PredefinedStyle.Header]
                });
            }
            writer.WriteEndElement(); // Row

            foreach (var record in data)
            {
                ct.ThrowIfCancellationRequested();
                writer.WriteStartElement(new Row());
                foreach (var col in columns)
                {
                    var value = record[col.FieldName];
                    if (value == DBNull.Value || value is null)
                    {
                        writer.WriteElement(new Cell());
                        continue;
                    }
                    var cell = CreateCell(value, col, styleMap);
                    writer.WriteElement(cell);
                }
                writer.WriteEndElement();
            }
            writer.WriteEndElement(); // SheetData

            if (options.AutoFilter)
            {
                var endCol = GetColumnName(columns.Count);
                writer.WriteElement(new AutoFilter { Reference = $"A1:{endCol}1" });
            }

            writer.WriteEndElement(); // Worksheet
            writer.Close();
        }

            var sheets = workbookPart.Workbook.AppendChild(new Sheets());
            sheets.Append(new Sheet
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
                cell.CellValue = new CellValue(Convert.ToString(value, CultureInfo.InvariantCulture));
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
                cell.InlineString = new InlineString(new Text(value.ToString()));
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
        var columnName = string.Empty;
        while (dividend > 0)
        {
            var modulo = (dividend - 1) % 26;
            columnName = Convert.ToChar(65 + modulo) + columnName;
            dividend = (dividend - modulo) / 26;
        }
        return columnName;
    }
}
