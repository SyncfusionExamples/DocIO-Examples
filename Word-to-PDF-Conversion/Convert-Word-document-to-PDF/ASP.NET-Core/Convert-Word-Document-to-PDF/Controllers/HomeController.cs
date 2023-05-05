using Convert_Word_Document_to_PDF.Models;
using Microsoft.AspNetCore.Mvc;
using System.Diagnostics;
using System.IO;
using Syncfusion.DocIO.DLS;
using Syncfusion.DocIO;
using Syncfusion.DocIORenderer;
using Syncfusion.Pdf;

namespace Convert_Word_Document_to_PDF.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;

        public HomeController(ILogger<HomeController> logger)
        {
            _logger = logger;
        }

        public IActionResult Index()
        {
            return View();
        }

        public ActionResult ConvertWordtoPDF()
        {
            //Open the file as Stream
            using (FileStream docStream = new FileStream(Path.GetFullPath("Data/Template.docx"), FileMode.Open, FileAccess.Read))
            {
                //Loads file stream into Word document
                using (WordDocument wordDocument = new WordDocument(docStream, Syncfusion.DocIO.FormatType.Automatic))
                {
                    //Instantiation of DocIORenderer for Word to PDF conversion
                    using (DocIORenderer render = new DocIORenderer())
                    {
                        //Converts Word document into PDF document
                        PdfDocument pdfDocument = render.ConvertToPDF(wordDocument);

                        //Saves the PDF document to MemoryStream.
                        MemoryStream stream = new MemoryStream();
                        pdfDocument.Save(stream);
                        stream.Position = 0;

                        //Download PDF document in the browser.
                        return File(stream, "application/pdf", "Sample.pdf");
                    }
                }
            }
        }
        public IActionResult Privacy()
        {
            return View();
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }
    }
}