using DataToExcel.Models;
using DataToExcel.Services;
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
        var response = provider.BuildStylesheet(out var map);

        // Then
        Assert.True(response.IsSuccess);
        Assert.NotNull(response.Data);
        Assert.Equal(1u, map[PredefinedStyle.Header]);
    }
}
