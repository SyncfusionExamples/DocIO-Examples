using Compare_Word_Documents.Models;
using Microsoft.AspNetCore.Mvc;
using System.Diagnostics;
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;

namespace Compare_Word_Documents.Controllers
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
        public ActionResult CompareWordDocuments()
        {

            //Loads the original document.
            using (FileStream originalDocumentStreamPath = new FileStream(Path.GetFullPath("Data/OriginalDocument.docx"), FileMode.Open, FileAccess.Read))
            {
                using (WordDocument originalDocument = new WordDocument(originalDocumentStreamPath, FormatType.Docx))
                {
                    //Loads the revised document.
                    using (FileStream revisedDocumentStreamPath = new FileStream(Path.GetFullPath("Data/RevisedDocument.docx"), FileMode.Open, FileAccess.Read))
                    {
                        using (WordDocument revisedDocument = new WordDocument(revisedDocumentStreamPath, FormatType.Docx))
                        {
                            //Compare original document with revised document.
                            originalDocument.Compare(revisedDocument, "Nancy Davolio", DateTime.Now.AddDays(-1));
                            // Create a memory stream to store the comparison result.
                            MemoryStream stream = new MemoryStream();

                            // Save the compared document into the MemoryStream.
                            originalDocument.Save(stream, FormatType.Docx);

                            //Download Word document in the browser.
                            return File(stream, "application/docx", "Result.docx");

                        }
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