using Convert_Word_document_to_PDF.Models;
using Microsoft.AspNetCore.Mvc;
using Syncfusion.DocIO.DLS;
using Syncfusion.DocIO;
using System.Diagnostics;
using Syncfusion.DocIORenderer;
using Syncfusion.Pdf;
using System.IO;
using Syncfusion.Drawing;

namespace Convert_Word_document_to_PDF.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;

        public HomeController(ILogger<HomeController> logger)
        {
            _logger = logger;
        }
        public ActionResult ConvertWordToPDF()
        {
            MemoryStream stream = new MemoryStream();
            try{
            using (FileStream inputStream = new FileStream(Path.GetFullPath("Data/Input.docx"), FileMode.Open, FileAccess.Read))
            {
                //Creating a new document.
                using (WordDocument document = new WordDocument(inputStream, FormatType.Docx))
                {                 
                    //Creates a new instance of DocIORenderer class.
                    using (DocIORenderer render = new DocIORenderer())
                    {
                        //Converts Word document into PDF document.
                        using (PdfDocument pdfDocument = render.ConvertToPDF(document))
                        {                          
                            //Saves the PDF document to MemoryStream.
                            pdfDocument.Save(stream);
                            stream.Position = 0;
                        }
                    }
                }
            }
            //Download PDF in the browser.
            return File(stream, "application/pdf", "WordtoPDF.pdf");
            }
            catch(Exception ex)
            {
               ViewBag.Message = ex.Message;
            }
            return View();
        }       
        public IActionResult Index()
        {
            return View();
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