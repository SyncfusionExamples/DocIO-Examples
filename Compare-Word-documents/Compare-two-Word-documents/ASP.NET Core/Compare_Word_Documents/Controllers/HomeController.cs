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
            string originalFilePath = Path.GetFullPath("Data/OriginalDocument.docx");
            string revisedFilePath = Path.GetFullPath("Data/RevisedDocument.docx");

            using (FileStream orgDocStream = new FileStream(originalFilePath, FileMode.Open, FileAccess.Read))
            using (FileStream revisedStream = new FileStream(revisedFilePath, FileMode.Open, FileAccess.Read))
            //Open the original Word document
            using (WordDocument originalDocument = new WordDocument(orgDocStream, FormatType.Docx))
            //Open the revised Word document
            using (WordDocument revisedDocument = new WordDocument(revisedStream, FormatType.Docx))
            {
                //Compare original document with revised document
                originalDocument.Compare(revisedDocument, "Andrew Fuller", DateTime.Now);
                // Create a memory stream to store the comparison result.
                MemoryStream stream = new MemoryStream();

                // Save the compared document into the MemoryStream.
                originalDocument.Save(stream, FormatType.Docx);

                //Download Word document in the browser.
                return File(stream, "application/docx", "Result.docx");
           
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