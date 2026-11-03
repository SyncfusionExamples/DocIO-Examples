using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.Functions.Worker;
using Microsoft.Extensions.Logging;
using Syncfusion.Licensing;
using WordtoPDF.Configuration;
using WordtoPDF.Services;

namespace SharePoint_To_Blob_Pdf_Converter;

public class Function1
{
    private readonly ILogger<Function1> _logger;

    public Function1(ILogger<Function1> logger)
    {
        _logger = logger;
    }

    [Function("ConvertDocuments")]
    public async Task<IActionResult> Run(
        [HttpTrigger(AuthorizationLevel.Function, "get", "post")] HttpRequest req)
    {
        _logger.LogInformation("Processing SharePoint documents to PDF...");

        var licenseKey = Environment.GetEnvironmentVariable("SYNCFUSION_LICENSE_KEY");
        if (!string.IsNullOrWhiteSpace(licenseKey))
            SyncfusionLicenseProvider.RegisterLicense(licenseKey);

        var azureAd = new AzureAdSettings();
        var sharePoint = new SharePointSettings();
        var blob = new BlobStorageSettings();

        var sharePointService = new SharePointService(azureAd, sharePoint);
        var converter = new PdfConverter();
        var blobService = new BlobStorageService(blob);

        int succeeded = 0, failed = 0;

        await foreach (var file in sharePointService.GetDocumentsAsync())
        {
            try
            {
                using var input = file.Content;
                using var pdf = converter.Convert(input, file.Name);
                var pdfName = Path.ChangeExtension(file.Name, ".pdf");
                var uri = await blobService.UploadPdfAsync(pdfName, pdf);
                _logger.LogInformation("OK    {Name} -> {Uri}", file.Name, uri);
                succeeded++;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "FAIL  {Name}: {Message}", file.Name, ex.Message);
                failed++;
            }
        }

        _logger.LogInformation("Done. Converted: {Succeeded}, Failed: {Failed}", succeeded, failed);
        return new OkObjectResult(new { Converted = succeeded, Failed = failed });
    }
}
