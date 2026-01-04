using Azure.Core;

namespace DataToExcel.Models;

public class ExcelExportRegistrationOptions
{
    public string ContainerName { get; set; } = string.Empty;
    public string? BlobPrefix { get; set; }
    public TimeSpan DefaultSasTtl { get; set; } = TimeSpan.FromHours(1);
    public string? ConnectionString { get; set; }
    public Uri? ContainerUri { get; set; }
    public TokenCredential? Credential { get; set; }

    public void Validate()
    {
        var hasConn = !string.IsNullOrWhiteSpace(ConnectionString);
        var hasUri = ContainerUri is not null;
        if (hasConn == hasUri)
            throw new InvalidOperationException("Specify either ConnectionString or ContainerUri but not both.");
        if (hasConn && string.IsNullOrWhiteSpace(ContainerName))
            throw new InvalidOperationException("ContainerName is required when using ConnectionString.");
    }
}
