using DataToExcel.Application;
using DataToExcel.Application.Interfaces;
using DataToExcel.Models;
using DataToExcel.Repositories;
using DataToExcel.Repositories.Interfaces;
using DataToExcel.Services;
using DataToExcel.Services.Interfaces;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;

namespace DataToExcel.Hosting;

public static class DependencyInjection
{
    public static IServiceCollection AddExcelExport(this IServiceCollection services, Action<ExcelExportRegistrationOptions> configure)
    {
        var options = new ExcelExportRegistrationOptions();
        configure(options);
        options.Validate();

        services.AddSingleton(options);
        services.AddTransient<IExcelStyleProvider, ExcelStyleProvider>();
        services.AddTransient<IExcelExportService, ExcelExportService>();
        services.AddTransient<IFileNamingService, FileNamingService>();
        services.AddSingleton<IBlobStorageRepository>(sp =>
        {
            if (!string.IsNullOrWhiteSpace(options.ConnectionString))
                return new AzureBlobStorageRepository(options.ConnectionString!, options.ContainerName, options.DefaultSasTtl);
            return new AzureBlobStorageRepository(options.ContainerUri!, options.Credential, options.DefaultSasTtl);
        });
        services.AddTransient<IExportExcel>(sp =>
            new ExportExcel(
                sp.GetRequiredService<IExcelExportService>(),
                sp.GetRequiredService<IFileNamingService>(),
                sp.GetRequiredService<IBlobStorageRepository>(),
                options));
        return services;
    }

    public static IServiceCollection AddExcelExport(this IServiceCollection services, IConfiguration configuration)
    {
        var bound = configuration.Get<ExcelExportRegistrationOptions>() ?? new ExcelExportRegistrationOptions();
        return services.AddExcelExport(opts =>
        {
            opts.ContainerName = bound.ContainerName;
            opts.BlobPrefix = bound.BlobPrefix;
            opts.ConnectionString = bound.ConnectionString;
            opts.ContainerUri = bound.ContainerUri;
            opts.Credential = bound.Credential;
            opts.DefaultSasTtl = bound.DefaultSasTtl;
        });
    }
}
