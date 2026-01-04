# DataToExcel

DataToExcel is a .NET 8 library that converts an `IEnumerable<IDataRecord>` into a streaming Excel file (`.xlsx`) and uploads it securely to Azure Blob Storage.

## Features
- Stream-based OpenXML writing for exports up to roughly 1 GB
- Predefined styles for currency, date, datetime, percentage, number, boolean, and text
- File names formatted as `<BaseName>_<DataDate:yyyyMMdd>_<Creation:yyyyMMdd_HHmmss>.xlsx`
- Secure upload to private containers with temporary read-only SAS links
- Works with dependency injection or as a standalone client

## Using dependency injection
1. Add the required NuGet packages to your project: `DocumentFormat.OpenXml`, `Azure.Storage.Blobs`, and `Azure.Identity`.
2. Register the exporter in your `IServiceCollection` by calling `AddExcelExport`.
3. Resolve `IExportExcel` and call `ExecuteAsync`.

```csharp
var services = new ServiceCollection();
services.AddExcelExport(options =>
{
    options.ConnectionString = "<SecureConnectionString>";
    options.ContainerName = "reports";
    options.DefaultSasTtl = TimeSpan.FromHours(2);
});

var provider = services.BuildServiceProvider();
var exporter = provider.GetRequiredService<IExportExcel>();

var result = await exporter.ExecuteAsync(
    data: records,
    columns: columns,
    baseFileName: "Sales",
    options: new ExcelExportOptions
    {
        SheetName = "Sales",
        DataDateUtc = new DateTime(2025, 8, 1),
        FreezeHeader = true,
        AutoFilter = true
    },
    ct: CancellationToken.None);
```

### Binding options from configuration
Store the connection string and other settings in `appsettings.json`, environment variables, or user secrets and bind them directly:

```json
{
  "ExcelExport": {
    "ConnectionString": "<SecureConnectionString>",
    "ContainerName": "reports",
    "BlobPrefix": "exports/finance",
    "DefaultSasTtl": "02:00:00"
  }
}
```

```csharp
var services = new ServiceCollection();
services.AddExcelExport(configuration.GetSection("ExcelExport"));
```

## Using the standalone client
When dependency injection is not available, instantiate `ExcelExportClient` directly. Two constructors are provided: one for connection strings and another for RBAC/AAD scenarios.

```csharp
// Using a connection string
var client = new ExcelExportClient("<SecureConnectionString>", "reports");

// Or using a container URI and optional credential (for RBAC/AAD)
var client = new ExcelExportClient(new Uri("https://account.blob.core.windows.net/reports"));

var result = await client.ExecuteAsync(
    data: records,
    columns: columns,
    baseFileName: "Sales",
    options: new ExcelExportOptions { SheetName = "Sales" },
    ct: CancellationToken.None);
```

The `ExecuteAsync` method returns a `BlobUploadResult` with the blob URI, SAS URI, and uploaded size.

## NuGet requirements
- DocumentFormat.OpenXml
- Azure.Storage.Blobs
- Azure.Identity (for RBAC/AAD scenarios)

## Limitations
- Maximum of 1,048,576 rows per worksheet
- `IAsyncEnumerable<IDataRecord>` and worksheet splitting are not yet implemented

## Performance and security notes
- Uses `OpenXmlWriter` with a temporary `FileStream` to keep memory usage low
- Files are uploaded to private containers and returned with a short-lived read-only SAS token
