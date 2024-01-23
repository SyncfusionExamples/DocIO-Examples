using Microsoft.AspNetCore.Mvc;
using Syncfusion.DocIO.DLS;
using Syncfusion.DocIORenderer;
using Syncfusion.Pdf;

namespace Word_To_PDF_Web_API.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class Word_To_PDF_Web_APIController : ControllerBase
    {
        private readonly IWebHostEnvironment _webHostEnvironment;

        public Word_To_PDF_Web_APIController(IWebHostEnvironment webHostEnvironment)
        {
            _webHostEnvironment = webHostEnvironment;
        }
        public string filePath;
        //[HttpPost("FileUpload")]

        private string FileUpload(IFormFile file)
        {

            string fileName = file.FileName;
            string directoryPath = Path.Combine(_webHostEnvironment.ContentRootPath, "uploadedFiles");
            filePath = Path.Combine(directoryPath, fileName);
            using (var fileStream = new FileStream(filePath, FileMode.Create))
            {
                file.CopyTo(fileStream);
            }
            return filePath;
        }

        private IActionResult WordToPdf(IFormFile file)
        {
            string filePath = FileUpload(file);
            //Loads file stream into Word document
            FileStream docStream = new FileStream(filePath, FileMode.Open, FileAccess.Read);
            //Instantiation of DocIORenderer for Word to PDF conversion
            WordDocument wordDocument = new WordDocument(docStream, Syncfusion.DocIO.FormatType.Automatic);
            DocIORenderer render = new DocIORenderer();
            render.Settings.ChartRenderingOptions.ImageFormat = Syncfusion.OfficeChart.ExportImageFormat.Jpeg;
            //Converts Word document into PDF document
            PdfDocument pdfDocument = render.ConvertToPDF(wordDocument);
            render.Dispose();
            wordDocument.Dispose();
            //Saves the PDF document to MemoryStream.
            MemoryStream stream = new MemoryStream();
            pdfDocument.Save(stream);
            stream.Position = 0;
            //Download PDF document in the browser.
            return File(stream, "application/pdf", "OutputFile.pdf");
        }


        // Convert Word to PDF
        [HttpPost("WordToPdf")]
        public IActionResult WordToPdfConversion(IFormFile file)
        {
            return WordToPdf(file);
        }

    }
}
