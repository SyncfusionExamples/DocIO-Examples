namespace WordtoPDF.Configuration;

public class AzureAdSettings
{
    public string TenantId { get; set; } = "<Enter your Tenant Id>";
    public string ClientId { get; set; } = "<Enter your Client Id>";
    public string ClientSecret { get; set; } = "<Enter your ClientSecret>";
}

public class SharePointSettings
{
    public string SiteUrl { get; set; } = "<Enter your Sharepoint URL>";
    public string LibraryName { get; set; } = "<Enter your LibraryName in your sharepoint>";
    public string FolderPath { get; set; } = "<Enter your FolderPath where the files are stored>";
}

public class BlobStorageSettings
{
    public string ConnectionString { get; set; } = "<Enter your Blob storage connection string>";
    public string ContainerName { get; set; } = "<Enter your Container name>";
}
