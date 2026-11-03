using Azure.Storage.Blobs;
using Azure.Storage.Blobs.Models;
using WordtoPDF.Configuration;

namespace WordtoPDF.Services;

public class BlobStorageService
{
    private readonly BlobContainerClient _container;

    public BlobStorageService(BlobStorageSettings settings)
    {
        var service = new BlobServiceClient(settings.ConnectionString);
        _container = service.GetBlobContainerClient(settings.ContainerName);
    }

    public async Task<Uri> UploadPdfAsync(string blobName, Stream content, CancellationToken ct = default)
    {
        await _container.CreateIfNotExistsAsync(cancellationToken: ct);

        var blob = _container.GetBlobClient(blobName);
        await blob.UploadAsync(content, new BlobUploadOptions
        {
            HttpHeaders = new BlobHttpHeaders { ContentType = "application/pdf" }
        }, ct);
        return blob.Uri;
    }
}
