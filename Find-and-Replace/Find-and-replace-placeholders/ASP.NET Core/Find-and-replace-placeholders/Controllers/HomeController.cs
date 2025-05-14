using System.Diagnostics;
using System.Text.RegularExpressions;
using Find_and_replace_placeholders.Models;
using Microsoft.AspNetCore.Mvc;
using Syncfusion.DocIO.DLS;
using Syncfusion.DocIO;
using Syncfusion.DocIORenderer;
using Syncfusion.Pdf;

namespace Find_and_replace_placeholders.Controllers
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

        public ActionResult CreateDocument(string firstname, string lastname, string email, string phone, string address)
        {
            //Creating a new document.
            using (WordDocument document = new WordDocument(new FileStream("Data/Template.docx", FileMode.Open, FileAccess.Read), FormatType.Automatic))
            {
                // Store inputs in a 1D array
                string[] userInputs = { firstname, lastname, email, phone, address };
                // Find all placeholders like "{{Name}}"
                TextSelection[] selections = document.FindAll(new Regex(@"\{(.*)\}"));
                for (int i = 0; i < selections.Count(); i++)
                {
                    TextSelection selection = selections[i];
                    //Replace the text with user values
                    document.Replace(selection.SelectedText, userInputs[i], false, true);
                }
                //Instantiation of DocIORenderer for Word to PDF conversion
                using (DocIORenderer render = new DocIORenderer())
                {
                    //Converts Word document into PDF document
                    PdfDocument pdfDocument = render.ConvertToPDF(document);

                    //Saves the PDF document to MemoryStream.
                    MemoryStream stream = new MemoryStream();
                    pdfDocument.Save(stream);
                    stream.Position = 0;

                    //Download PDF document in the browser.
                    return File(stream, "application/pdf", "Sample.pdf");
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
