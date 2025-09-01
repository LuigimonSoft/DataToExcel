using DataToExcel.Models;

namespace DataToExcel.Models;

public record ColumnDefinition(
    string FieldName,
    string Title,
    ColumnDataType DataType,
    double? Width = null,
    PredefinedStyle? Style = null,
    string? NumberFormatCode = null);
