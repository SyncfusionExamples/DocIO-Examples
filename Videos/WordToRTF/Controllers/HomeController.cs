using System.Diagnostics;
using Microsoft.AspNetCore.Mvc;
using WordToRTF.Models;
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;

namespace WordToRTF.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;

        public HomeController(ILogger<HomeController> logger)
        {
            _logger = logger;
        }

        public IActionResult WordToRTF()
        {
            // Open the Word document as a file stream
            FileStream fileStream = new FileStream("Data\\Template.docx", FileMode.Open, FileAccess.Read, FileShare.ReadWrite);

            // Load the stream into a DocIO WordDocument instance
            WordDocument document = new WordDocument(fileStream, FormatType.Docx);

            // Create a memory stream to store the converted RTF content
            MemoryStream outputStream = new MemoryStream();

            // Save the document in RTF format
            document.Save(outputStream, FormatType.Rtf);

            // Close the document to release resources
            document.Close();

            // Return the downloadable file
            return File(outputStream, "application/rtf", "WordToRTF.rtf");
        }

        public IActionResult RTFToWord()
        {
            // Open the RTF file as a file stream
            FileStream fileStream = new FileStream("Data\\Input.rtf", FileMode.Open, FileAccess.Read, FileShare.ReadWrite);

            // Load the stream into a DocIO WordDocument instance
            WordDocument document = new WordDocument(fileStream, FormatType.Rtf);

            // Create a memory stream to store the converted Word content
            MemoryStream outputStream = new MemoryStream();

            // Save the document in DOCX format
            document.Save(outputStream, FormatType.Docx);

            // Close the document to release resources
            document.Close();

            // Return the downloadable file
            return File(outputStream, "application/docx", "RTFToWord.docx");
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
