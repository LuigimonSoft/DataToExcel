using DataToExcel.Services;
using Xunit;

namespace DataToExcel.Test.Services;

public class FileNamingServiceTests
{
    [Fact]
    public void GivenRawNameWhenComposeExcelFileNameThenResultShouldBeSanitizedAndFormatted()
    {
        // Given
        var svc = new FileNamingService();

        // When
        var response = svc.ComposeExcelFileName("Test/Report", new DateTime(2024,1,2), new DateTime(2024,1,3,12,5,6));

        // Then
        Assert.True(response.IsSuccess);
        Assert.Equal("Test_Report_20240102_20240103_120506.xlsx", response.Data);
    }
}
