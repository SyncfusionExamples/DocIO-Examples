using Microsoft.AspNetCore.Mvc;
using Syncfusion.DocIO.DLS;
using Syncfusion.DocIORenderer;
using Syncfusion.Pdf;

namespace Convert_Word_Document_to_PDF.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class ValuesController : ControllerBase
    {
        [HttpGet]
        [Route("api/ConvertWordToPdf")]
        public IActionResult ConvertWordToPdf()
        {
            try
            {
                var fileDownloadName = "Output.pdf";
                const string contentType = "application/pdf";
                var stream = ConvertWordDocumentToPdf();
                stream.Position = 0;
                return File(stream, contentType, fileDownloadName);
            }
            catch (Exception ex)
            {
                return BadRequest("Error occurred while converting Word to PDF: " + ex.Message);
            }
        }

        public static MemoryStream ConvertWordDocumentToPdf()
        {
            //Open the existing PowerPoint presentation with loaded stream.
            using (WordDocument wordDocument = new WordDocument(Path.GetFullPath("Data/Template.docx")))
            {
                using (DocIORenderer render = new DocIORenderer())
                {
                    PdfDocument pdfDocument = render.ConvertToPDF(wordDocument);
                    //Create the MemoryStream to save the converted PDF.      
                    MemoryStream pdfStream = new MemoryStream();
                    //Save the converted PDF document to MemoryStream.
                    pdfDocument.Save(pdfStream);
                    pdfStream.Position = 0;
                    //Download PDF document in the browser.
                    return pdfStream;
                }
            }
        }
    }
}
