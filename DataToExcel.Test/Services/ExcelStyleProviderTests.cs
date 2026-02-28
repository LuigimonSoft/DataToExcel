using DataToExcel.Models;
using DataToExcel.Services;
using DocumentFormat.OpenXml.Spreadsheet;
using Xunit;

namespace DataToExcel.Test.Services;

public class ExcelStyleProviderTests
{
    [Fact]
    public void GivenStyleProviderWhenBuildStylesheetThenMapShouldContainHeaderStyle()
    {
        // Given
        var provider = new ExcelStyleProvider();

        // When
        var response = provider.BuildStylesheet(new ExcelExportOptions(), out var map);

        // Then
        Assert.True(response.IsSuccess);
        Assert.NotNull(response.Data);
        Assert.Equal(1u, map[PredefinedStyle.Header]);
    }

    [Fact]
    public void GivenHeaderBackgroundColorWhenBuildStylesheetThenHeaderFillShouldBeConfigured()
    {
        // Given
        var provider = new ExcelStyleProvider();

        // When
        var response = provider.BuildStylesheet(new ExcelExportOptions { HeaderBackgroundColorHex = "#00FF00" }, out _);

        // Then
        Assert.True(response.IsSuccess);
        var stylesheet = Assert.IsType<Stylesheet>(response.Data);
        var headerFormat = stylesheet.CellFormats!.Elements<CellFormat>().ElementAt(1);
        Assert.True(headerFormat.ApplyFill?.Value);

        var headerFill = stylesheet.Fills!.Elements<Fill>().ElementAt((int)headerFormat.FillId!.Value);
        var patternFill = Assert.IsType<PatternFill>(headerFill.FirstChild);
        Assert.Equal(PatternValues.Solid, patternFill.PatternType!.Value);
        Assert.Equal("00FF00", patternFill.ForegroundColor!.Rgb!.Value);
    }

    [Fact]
    public void GivenHeaderTextColorWhenBuildStylesheetThenHeaderFontColorShouldBeConfigured()
    {
        // Given
        var provider = new ExcelStyleProvider();

        // When
        var response = provider.BuildStylesheet(new ExcelExportOptions { HeaderTextColorHex = "112233" }, out _);

        // Then
        Assert.True(response.IsSuccess);
        var stylesheet = Assert.IsType<Stylesheet>(response.Data);
        var headerFont = stylesheet.Fonts!.Elements<Font>().ElementAt(1);
        Assert.Equal("112233", headerFont.GetFirstChild<Color>()!.Rgb!.Value);
    }

    [Fact]
    public void GivenInvalidColorsWhenBuildStylesheetThenHeaderStyleShouldFallbackToDefaults()
    {
        // Given
        var provider = new ExcelStyleProvider();

        // When
        var response = provider.BuildStylesheet(new ExcelExportOptions
        {
            HeaderBackgroundColorHex = "BAD",
            HeaderTextColorHex = "NOT_A_COLOR"
        }, out _);

        // Then
        Assert.True(response.IsSuccess);
        var stylesheet = Assert.IsType<Stylesheet>(response.Data);
        var headerFormat = stylesheet.CellFormats!.Elements<CellFormat>().ElementAt(1);
        Assert.False(headerFormat.ApplyFill?.Value ?? false);

        var headerFont = stylesheet.Fonts!.Elements<Font>().ElementAt(1);
        Assert.Null(headerFont.GetFirstChild<Color>());
    }

    [Theory]
    [InlineData("", "")]
    [InlineData("   ", "\t")]
    [InlineData("", "NOT_VALID")]
    [InlineData("BAD", "")]
    public void GivenBlankOrMixedInvalidHeaderColorsWhenBuildStylesheetThenHeaderStyleShouldFallbackToDefaults(
        string? backgroundColor,
        string? textColor)
    {
        // Given
        var provider = new ExcelStyleProvider();

        // When
        var response = provider.BuildStylesheet(new ExcelExportOptions
        {
            HeaderBackgroundColorHex = backgroundColor,
            HeaderTextColorHex = textColor
        }, out _);

        // Then
        Assert.True(response.IsSuccess);
        var stylesheet = Assert.IsType<Stylesheet>(response.Data);

        var headerFormat = stylesheet.CellFormats!.Elements<CellFormat>().ElementAt(1);
        Assert.False(headerFormat.ApplyFill?.Value ?? false);

        var headerFont = stylesheet.Fonts!.Elements<Font>().ElementAt(1);
        Assert.Null(headerFont.GetFirstChild<Color>());
    }
}
