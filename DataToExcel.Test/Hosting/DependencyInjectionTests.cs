using DataToExcel.Hosting;
using DataToExcel.Models;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Xunit;

namespace DataToExcel.Test.Hosting;

public class DependencyInjectionTests
{
    [Fact]
    public void GivenConfigurationWhenAddExcelExportThenOptionsShouldBind()
    {
        // Given
        var settings = new Dictionary<string, string?>
        {
            ["ConnectionString"] = "UseDevelopmentStorage=true",
            ["ContainerName"] = "reports",
            ["BlobPrefix"] = "exports/finance"
        };
        var configuration = new ConfigurationBuilder()
            .AddInMemoryCollection(settings)
            .Build();

        // When
        var services = new ServiceCollection();
        services.AddExcelExport(configuration);
        var provider = services.BuildServiceProvider();

        // Then
        var opts = provider.GetRequiredService<ExcelExportRegistrationOptions>();
        Assert.Equal("UseDevelopmentStorage=true", opts.ConnectionString);
        Assert.Equal("reports", opts.ContainerName);
        Assert.Equal("exports/finance", opts.BlobPrefix);
    }
}
