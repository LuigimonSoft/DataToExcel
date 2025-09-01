using DataToExcel.Models;
using DataToExcel.Services.Interfaces;
using DocumentFormat.OpenXml.Spreadsheet;

namespace DataToExcel.Services;

public class ExcelStyleProvider : IExcelStyleProvider
{
    public ServiceResponse<Stylesheet> BuildStylesheet(out IReadOnlyDictionary<PredefinedStyle, uint> styleIndexMap)
    {
        try
        {
            var fonts = new Fonts(
                new Font(),
                new Font(new Bold())
            );
            var fills = new Fills(
                new Fill(new PatternFill { PatternType = PatternValues.None }),
                new Fill(new PatternFill { PatternType = PatternValues.Gray125 })
            );
            var borders = new Borders(new Border());
            var cellStyleFormats = new CellStyleFormats(new CellFormat());

            uint nfId = 164; // custom formats
            var numberingFormats = new NumberingFormats();
            var cellFormats = new List<CellFormat>
            {
                new(),                                // 0 default
                new() { FontId = 1, ApplyFont = true } // 1 header
            };

            // number
            numberingFormats.Append(new NumberingFormat { NumberFormatId = nfId, FormatCode = "#,##0.00" });
            cellFormats.Add(new() { NumberFormatId = nfId, ApplyNumberFormat = true });
            var numberIdx = (uint)cellFormats.Count - 1;
            nfId++;

            // date
            numberingFormats.Append(new NumberingFormat { NumberFormatId = nfId, FormatCode = "yyyy-mm-dd" });
            cellFormats.Add(new() { NumberFormatId = nfId, ApplyNumberFormat = true });
            var dateIdx = (uint)cellFormats.Count - 1;
            nfId++;

            // datetime
            numberingFormats.Append(new NumberingFormat { NumberFormatId = nfId, FormatCode = "yyyy-mm-dd hh:mm:ss" });
            cellFormats.Add(new() { NumberFormatId = nfId, ApplyNumberFormat = true });
            var dateTimeIdx = (uint)cellFormats.Count - 1;
            nfId++;

            // currency
            numberingFormats.Append(new NumberingFormat { NumberFormatId = nfId, FormatCode = "#,##0.00" });
            cellFormats.Add(new() { NumberFormatId = nfId, ApplyNumberFormat = true });
            var currencyIdx = (uint)cellFormats.Count - 1;
            nfId++;

            // percentage
            numberingFormats.Append(new NumberingFormat { NumberFormatId = nfId, FormatCode = "0.00%" });
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

            var stylesheet = new Stylesheet(numberingFormats, fonts, fills, borders, cellStyleFormats, new CellFormats(cellFormats));
            return new ServiceResponse<Stylesheet>(stylesheet) { IsSuccess = true };
        }
        catch (Exception ex)
        {
            styleIndexMap = new Dictionary<PredefinedStyle, uint>();
            return new ServiceResponse<Stylesheet> { IsSuccess = false, ErrorMessage = ex.Message };
        }
    }
}
