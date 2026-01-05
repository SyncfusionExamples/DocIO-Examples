using Microsoft.AspNetCore.Mvc;
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.Diagnostics;
using WordToText.Models;

namespace WordToText.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;

        public HomeController(ILogger<HomeController> logger)
        {
            _logger = logger;
        }

        public IActionResult WordToText()
        {
            // Loads an existing Word document into DocIO instance
            FileStream fileStreamPath = new FileStream("Data\\Template.docx", FileMode.Open, FileAccess.Read, FileShare.ReadWrite);

            // Load the document into DocIO
            WordDocument document = new WordDocument(fileStreamPath, FormatType.Docx);

            // Create an output memory stream to hold the document content
            MemoryStream outputStream = new MemoryStream();

            // Save the new document to the output stream
            document.Save(outputStream, FormatType.Txt);

            // Close the document 
            document.Close();

            //Return the output file
            return File(outputStream, "application/txt", "WordToText.txt");
        }

        public IActionResult TextToWord()
        {
            //Loads an existing Word document into DocIO instance
            FileStream fileStreamPath = new FileStream("Data\\Template.txt", FileMode.Open, FileAccess.Read, FileShare.ReadWrite);

            // Load the document into DocIO
            WordDocument document = new WordDocument(fileStreamPath, FormatType.Txt);

            //Create an output memory stream to hold the document content
            MemoryStream outputStream = new MemoryStream();

            // Save the new document to the output stream
            document.Save(outputStream, FormatType.Docx);

            //Close the document 
            document.Close();

            //Return the output file
            return File(outputStream, "application/docx", "TextToWord.docx");
        }

        public IActionResult ExtractPlainText()
        {
            //Loads an existing Word document into DocIO instance
            FileStream fileStreamPath = new FileStream("Data\\Template.docx", FileMode.Open, FileAccess.Read, FileShare.ReadWrite);

            // Load the document into DocIO
            WordDocument document = new WordDocument(fileStreamPath, FormatType.Docx);

            //Get the plain text from the document
            string text = document.GetText();

            //Create a new Word document to hold the extracted text
            WordDocument newdocument = new WordDocument();

            //Adds new section
            IWSection section = newdocument.AddSection();

            //Adds new paragraph
            IWParagraph paragraph = section.AddParagraph();

            //Append the extracted text to the new document.
            paragraph.AppendText(text);

            //Create an output stream for the new document
            MemoryStream outputStream = new MemoryStream();

            //Save the new document to the output stream
            newdocument.Save(outputStream, FormatType.Docx);

            //Close the document 
            newdocument.Close();

            //Return the output file
            return File(outputStream, "application/docx", "ExtractPlainText.docx");
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
