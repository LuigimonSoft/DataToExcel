using System.Data;
using Azure.Core;
using DataToExcel.Application;
using DataToExcel.Application.Interfaces;
using DataToExcel.Models;
using DataToExcel.Repositories;
using DataToExcel.Repositories.Interfaces;
using DataToExcel.Services;
using DataToExcel.Services.Interfaces;
using DataToExcel.Wrappers.Interfaces;

namespace DataToExcel;

public class ExcelExportClient : IExportExcel
{
    private readonly ExportExcel _inner;

    public ExcelExportClient(string connectionString, string containerName, TimeSpan? defaultSasTtl = null, string? blobPrefix = null)
    {
        var options = new ExcelExportRegistrationOptions
        {
            ConnectionString = connectionString,
            ContainerName = containerName,
            DefaultSasTtl = defaultSasTtl ?? TimeSpan.FromHours(1),
            BlobPrefix = blobPrefix
        };
        _inner = Build(options);
    }

    public ExcelExportClient(Uri containerUri, TokenCredential? credential = null, TimeSpan? defaultSasTtl = null, string? blobPrefix = null)
    {
        var options = new ExcelExportRegistrationOptions
        {
            ContainerUri = containerUri,
            Credential = credential,
            DefaultSasTtl = defaultSasTtl ?? TimeSpan.FromHours(1),
            BlobPrefix = blobPrefix
        };
        _inner = Build(options);
    }

    public ExcelExportClient(IBlobContainerClient container, TimeSpan? defaultSasTtl = null, string? blobPrefix = null)
    {
        var options = new ExcelExportRegistrationOptions
        {
            ContainerName = container.Name,
            DefaultSasTtl = defaultSasTtl ?? TimeSpan.FromHours(1),
            BlobPrefix = blobPrefix
        };
        _inner = Build(options, container);
    }

    private static ExportExcel Build(ExcelExportRegistrationOptions options, IBlobContainerClient? container = null)
    {
        if (container is null)
            options.Validate();
        IExcelStyleProvider style = new ExcelStyleProvider();
        IExcelExportService export = new ExcelExportService(style);
        IFileNamingService naming = new FileNamingService();
        IBlobStorageRepository repo;
        if (container is null)
        {
            if (!string.IsNullOrWhiteSpace(options.ConnectionString))
            {
                repo = new AzureBlobStorageRepository(options.ConnectionString!, options.ContainerName, options.DefaultSasTtl);
            }
            else
            {
                repo = new AzureBlobStorageRepository(options.ContainerUri!, options.Credential, options.DefaultSasTtl);
            }
        }
        else
        {
            repo = new AzureBlobStorageRepository(container, options.DefaultSasTtl);
        }
        return new ExportExcel(export, naming, repo, options);
    }

    public async Task<BlobUploadResult> ExecuteAsync(IEnumerable<IDataRecord> data,
        IReadOnlyList<ColumnDefinition> columns,
        string baseFileName,
        ExcelExportOptions options,
        TimeSpan? sasTtl = null,
        CancellationToken ct = default) =>
        await _inner.ExecuteAsync(data, columns, baseFileName, options, sasTtl, ct);
}
