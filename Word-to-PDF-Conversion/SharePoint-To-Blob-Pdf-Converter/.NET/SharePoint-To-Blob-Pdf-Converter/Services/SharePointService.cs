using Azure.Identity;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using WordtoPDF.Configuration;

namespace WordtoPDF.Services;

public record SharePointFile(string Name, MemoryStream Content);

public class SharePointService
{
    private readonly GraphServiceClient _graph;
    private readonly SharePointSettings _settings;

    public SharePointService(AzureAdSettings azureAd, SharePointSettings settings)
    {
        _settings = settings;
        var credential = new ClientSecretCredential(azureAd.TenantId, azureAd.ClientId, azureAd.ClientSecret);
        _graph = new GraphServiceClient(credential, new[] { "https://graph.microsoft.com/.default" });
    }

    public async IAsyncEnumerable<SharePointFile> GetDocumentsAsync(
        [System.Runtime.CompilerServices.EnumeratorCancellation] CancellationToken ct = default)
    {
        var siteUri = new Uri(_settings.SiteUrl);
        string siteID = $"{siteUri.Host}:{siteUri.AbsolutePath.TrimEnd('/')}";
        var site = await _graph.Sites[siteID]
            .GetAsync(cancellationToken: ct)
            ?? throw new InvalidOperationException($"SharePoint site not found: {_settings.SiteUrl}");

        var drives = await _graph.Sites[site.Id].Drives.GetAsync(cancellationToken: ct);
        var drive = drives?.Value?.FirstOrDefault(d =>
                string.Equals(d.Name, _settings.LibraryName, StringComparison.OrdinalIgnoreCase))
            ?? throw new InvalidOperationException($"Document library '{_settings.LibraryName}' not found.");

        var folderPath = _settings.FolderPath.Trim('/');
        var itemId = "root";
        if (folderPath.Length > 0)
        {
            var folder = await _graph.Drives[drive.Id].Root.ItemWithPath(folderPath)
                .GetAsync(cancellationToken: ct)
                ?? throw new InvalidOperationException($"Folder '{folderPath}' not found.");
            itemId = folder.Id!;
        }

        var page = await _graph.Drives[drive.Id].Items[itemId].Children.GetAsync(cancellationToken: ct);
        while (page != null)
        {
            foreach (var item in page.Value ?? new List<DriveItem>())
            {
                if (item.File == null || item.Name == null ||
                    !(item.Name.EndsWith(".docx", StringComparison.OrdinalIgnoreCase) ||
                      item.Name.EndsWith(".pptx", StringComparison.OrdinalIgnoreCase)) ||
                    item.Name.StartsWith("~$"))
                    continue;

                await using var stream = await _graph.Drives[drive.Id].Items[item.Id].Content
                    .GetAsync(cancellationToken: ct);
                if (stream == null) continue;

                var buffer = new MemoryStream();
                await stream.CopyToAsync(buffer, ct);
                buffer.Position = 0;
                yield return new SharePointFile(item.Name, buffer);
            }

            if (page.OdataNextLink == null) break;
            page = await _graph.Drives[drive.Id].Items[itemId].Children
                .WithUrl(page.OdataNextLink).GetAsync(cancellationToken: ct);
        }
    }
}
