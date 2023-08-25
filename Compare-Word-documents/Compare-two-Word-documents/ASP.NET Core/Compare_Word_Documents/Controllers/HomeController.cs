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
            //Open the Original file as Stream.
            using (FileStream originalDocStream = new FileStream(Path.GetFullPath("Data/OriginalDocument.docx"), FileMode.Open, FileAccess.Read))
            {
                //Open the Revised file as Stream.
                using (FileStream revisedDocStream = new FileStream(Path.GetFullPath("Data/RevisedDocument.docx"), FileMode.Open, FileAccess.Read))
                {
                    //Loads Original file stream into Word document.
                    using (WordDocument originalWordDocument = new WordDocument(originalDocStream, FormatType.Docx))
                    {
                        //Loads Revised file stream into Word document.
                        using (WordDocument revisedWordDocument = new WordDocument(revisedDocStream, FormatType.Docx))
                        {
                            // Create a memory stream to store the comparison result.
                            MemoryStream stream = new MemoryStream();

                            // Compare the original and revised Word documents.
                            originalWordDocument.Compare(revisedWordDocument, "Your Name", DateTime.Now);

                            // Save the compared document into the MemoryStream.
                            originalWordDocument.Save(stream,FormatType.Docx);

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