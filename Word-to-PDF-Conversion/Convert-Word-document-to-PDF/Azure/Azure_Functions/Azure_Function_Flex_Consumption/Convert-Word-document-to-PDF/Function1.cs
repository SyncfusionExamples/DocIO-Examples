using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.Functions.Worker;
using Microsoft.Extensions.Logging;
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.DocIORenderer;
using Syncfusion.Pdf;

namespace Convert_Word_document_to_PDF;

public class Function1
{
    private readonly ILogger<Function1> _logger;

    public Function1(ILogger<Function1> logger)
    {
        _logger = logger;
    }

    [Function("ConvertWordtoPDF")]
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
            using WordDocument document = new WordDocument(inputStream, FormatType.Automatic);
            // Attach font substitution handler to manage missing fonts
            document.FontSettings.SubstituteFont += FontSettings_SubstituteFont;
            // Initialize DocIORenderer to convert Word document to PDF
            using DocIORenderer renderer = new DocIORenderer();
            // Convert the Word document to a PDF document
            using PdfDocument pdfDoc = renderer.ConvertToPDF(document);
            // Create a memory stream to hold the PDF output
            await using MemoryStream outputStream = new MemoryStream();
            // Save the PDF into the output stream
            pdfDoc.Save(outputStream);
            // Close the PDF document and release resources
            pdfDoc.Close(true);
            // Reset stream position to the beginning for reading
            outputStream.Position = 0;
            // Convert the PDF stream to a byte array
            var pdfBytes = outputStream.ToArray();

            // Create a file result to return the PDF as a downloadable file
            var fileResult = new FileContentResult(pdfBytes, "application/pdf")
            {
                FileDownloadName = "converted.pdf"
            };
            // Return the PDF file result to the client
            return fileResult;
        }
        catch (Exception ex)
        {
            // Log the error with details for troubleshooting
            _logger.LogError(ex, "Error converting DOCX to PDF.");
            // Prepare error message including exception details
            var msg = $"Exception: {ex.Message}\n\n{ex}";
            // Return a 500 Internal Server Error response with the message
            return new ContentResult { StatusCode = 500, Content = msg, ContentType = "text/plain; charset=utf-8" };
        }
    }
    /// <summary>
    /// Event handler for font substitution during PDF conversion
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