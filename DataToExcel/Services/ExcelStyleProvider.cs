using DataToExcel.Models;
using DataToExcel.Services.Interfaces;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace DataToExcel.Services;

public class ExcelStyleProvider : IExcelStyleProvider
{
    public ServiceResponse<Stylesheet> BuildStylesheet(ExcelExportOptions options, out IReadOnlyDictionary<PredefinedStyle, uint> styleIndexMap)
    {
        try
        {
            var fonts = new Fonts();
            fonts.AppendChild(new Font());

            var headerFont = new Font();
            headerFont.AppendChild(new Bold());
            var headerTextColor = NormalizeHexColor(options.HeaderTextColorHex);
            if (headerTextColor is not null)
            {
                headerFont.AppendChild(new Color { Rgb = headerTextColor });
            }
            fonts.AppendChild(headerFont);

            var fills = new Fills();
            var fillNone = new Fill();
            fillNone.AppendChild(new PatternFill { PatternType = PatternValues.None });
            fills.AppendChild(fillNone);
            var fillGray = new Fill();
            fillGray.AppendChild(new PatternFill { PatternType = PatternValues.Gray125 });
            fills.AppendChild(fillGray);

            uint headerFillId = 0;
            var headerBackgroundColor = NormalizeHexColor(options.HeaderBackgroundColorHex);
            if (headerBackgroundColor is not null)
            {
                var headerFill = new Fill();
                headerFill.AppendChild(new PatternFill
                {
                    PatternType = PatternValues.Solid,
                    ForegroundColor = new ForegroundColor { Rgb = headerBackgroundColor },
                    BackgroundColor = new BackgroundColor { Indexed = 64 }
                });
                fills.AppendChild(headerFill);
                headerFillId = (uint)fills.Count() - 1;
            }

            var borders = new Borders();
            borders.AppendChild(new Border());
            var cellStyleFormats = new CellStyleFormats();
            cellStyleFormats.AppendChild(new CellFormat());

            uint nfId = 164; // custom formats
            var numberingFormats = new NumberingFormats();
            var cellFormats = new List<CellFormat>
            {
                new(),                                // 0 default
                new() { FontId = 1, FillId = headerFillId, ApplyFont = true, ApplyFill = headerBackgroundColor is not null } // 1 header
            };

            // number
            numberingFormats.AppendChild(new NumberingFormat { NumberFormatId = nfId, FormatCode = "#,##0.00" });
            cellFormats.Add(new() { NumberFormatId = nfId, ApplyNumberFormat = true });
            var numberIdx = (uint)cellFormats.Count - 1;
            nfId++;

            // date
            numberingFormats.AppendChild(new NumberingFormat { NumberFormatId = nfId, FormatCode = "yyyy-mm-dd" });
            cellFormats.Add(new() { NumberFormatId = nfId, ApplyNumberFormat = true });
            var dateIdx = (uint)cellFormats.Count - 1;
            nfId++;

            // datetime
            numberingFormats.AppendChild(new NumberingFormat { NumberFormatId = nfId, FormatCode = "yyyy-mm-dd hh:mm:ss" });
            cellFormats.Add(new() { NumberFormatId = nfId, ApplyNumberFormat = true });
            var dateTimeIdx = (uint)cellFormats.Count - 1;
            nfId++;

            // currency
            numberingFormats.AppendChild(new NumberingFormat { NumberFormatId = nfId, FormatCode = "#,##0.00" });
            cellFormats.Add(new() { NumberFormatId = nfId, ApplyNumberFormat = true });
            var currencyIdx = (uint)cellFormats.Count - 1;
            nfId++;

            // percentage
            numberingFormats.AppendChild(new NumberingFormat { NumberFormatId = nfId, FormatCode = "0.00%" });
            cellFormats.Add(new() { NumberFormatId = nfId, ApplyNumberFormat = true });
            var percentageIdx = (uint)cellFormats.Count - 1;

            // boolean
            cellFormats.Add(new());
            var boolIdx = (uint)cellFormats.Count - 1;

            // text
            cellFormats.Add(new() { });
            var textIdx = (uint)cellFormats.Count - 1;

            styleIndexMap = new Dictionary<PredefinedStyle, uint>
            {
                [PredefinedStyle.Default] = 0,
                [PredefinedStyle.Header] = 1,
                [PredefinedStyle.Number] = numberIdx,
                [PredefinedStyle.Date] = dateIdx,
                [PredefinedStyle.DateTime] = dateTimeIdx,
                [PredefinedStyle.Currency] = currencyIdx,
                [PredefinedStyle.Percentage] = percentageIdx,
                [PredefinedStyle.Boolean] = boolIdx,
                [PredefinedStyle.Text] = textIdx
            };

            var cellFormatsElement = new CellFormats();
            foreach (var cf in cellFormats)
                cellFormatsElement.AppendChild(cf);

            var stylesheet = new Stylesheet();
            stylesheet.AppendChild(numberingFormats);
            stylesheet.AppendChild(fonts);
            stylesheet.AppendChild(fills);
            stylesheet.AppendChild(borders);
            stylesheet.AppendChild(cellStyleFormats);
            stylesheet.AppendChild(cellFormatsElement);
            return new ServiceResponse<Stylesheet>(stylesheet) { IsSuccess = true };
        }
        catch (Exception ex)
        {
            styleIndexMap = new Dictionary<PredefinedStyle, uint>();
            return new ServiceResponse<Stylesheet> { IsSuccess = false, ErrorMessage = ex.Message };
        }
    }

    private static HexBinaryValue? NormalizeHexColor(string? color)
    {
        if (string.IsNullOrWhiteSpace(color))
            return null;

        var normalized = color.Trim().TrimStart('#').ToUpperInvariant();
        if (normalized.Length is not 6 and not 8)
            return null;

        return normalized.All(Uri.IsHexDigit) ? normalized : null;
    }
}
