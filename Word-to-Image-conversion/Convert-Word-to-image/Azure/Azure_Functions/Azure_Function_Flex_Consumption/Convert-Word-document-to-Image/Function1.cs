using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.Functions.Worker;
using Microsoft.Extensions.Logging;
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.DocIORenderer;

namespace Convert_Word_document_to_Image;

public class Function1
{
    private readonly ILogger<Function1> _logger;

    public Function1(ILogger<Function1> logger)
    {
        _logger = logger;
    }

    [Function("ConvertWordDocumenttoImage")]
    public async Task<IActionResult> Run([HttpTrigger(AuthorizationLevel.Function, "post")] HttpRequest req)
    {
        try
        {
            // Create a memory stream to hold the incoming request body (Word document bytes)
            await using MemoryStream inputStream = new MemoryStream();
            // Copy the request body into the memory stream
            await req.Body.CopyToAsync(inputStream);
            // Check if the stream is empty (no file content received)
            if (inputStream.Length == 0)
                return new BadRequestObjectResult("No file content received in request body.");
            // Reset stream position to the beginning for reading
            inputStream.Position = 0;
            // Load the Word document from the stream (auto-detects format type)
            using WordDocument document = new WordDocument(inputStream, Syncfusion.DocIO.FormatType.Automatic);
            // Attach font substitution handler to manage missing fonts
            document.FontSettings.SubstituteFont += FontSettings_SubstituteFont;
            // Initialize the DocIORenderer to perform image conversion.
            DocIORenderer render = new DocIORenderer();
            // Convert Word document to image as stream.
            Stream imageStream = document.RenderAsImages(0, ExportImageFormat.Png);
            // Reset the stream position.
            imageStream.Position = 0;
            // Create a memory stream to hold the Image output
            await using MemoryStream outputStream = new MemoryStream();
            // Copy the contents of the image stream to the memory stream.
            await imageStream.CopyToAsync(outputStream);
            // Convert the Image stream to a byte array
            var imageBytes = outputStream.ToArray();
            //Reset the stream position.
            imageStream.Position = 0;
            // Reset stream position to the beginning for reading
            outputStream.Position = 0;
            // Create a file result to return the PNG as a downloadable file
            return new FileContentResult(imageBytes, "image/png")
            {
                FileDownloadName = "document-1.png"
            };
        }
        catch (Exception ex)
        {
            // Log the error with details for troubleshooting
            _logger.LogError(ex, "Error converting Word document to Image.");
            // Prepare error message including exception details
            var msg = $"Exception: {ex.Message}\n\n{ex}";
            // Return a 500 Internal Server Error response with the message
            return new ContentResult { StatusCode = 500, Content = msg, ContentType = "text/plain; charset=utf-8" };
        }
    }
    /// <summary>
    /// Event handler for font substitution during Image conversion
    /// </summary>
    /// <param name="sender"></param>
    /// <param name="args"></param>
    private void FontSettings_SubstituteFont(object sender, SubstituteFontEventArgs args)
    {
        // Define the path to the Fonts folder in the application base directory
        string fontsFolder = Path.Combine(AppContext.BaseDirectory, "Fonts");
        // If the original font is Calibri, substitute with calibri-regular.ttf
        if (args.OriginalFontName == "Calibri")
        {
            args.AlternateFontStream = File.OpenRead(Path.Combine(fontsFolder, "calibri-regular.ttf"));
        }
        // Otherwise, substitute with Times New Roman
        else
        {
            args.AlternateFontStream = File.OpenRead(Path.Combine(fontsFolder, "Times New Roman.ttf"));
        }
    }
}