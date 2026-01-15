using System.Diagnostics;
using Fields.Models;
using Microsoft.AspNetCore.Mvc;
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;

namespace Fields.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;

        public HomeController(ILogger<HomeController> logger)
        {
            _logger = logger;
        }

        public IActionResult AddField()
        {
            using (FileStream fileStream = new FileStream("Data\\Template.docx", FileMode.Open, FileAccess.Read))
            {
                WordDocument document = new WordDocument(fileStream, FormatType.Automatic);

                foreach (WSection section in document.Sections)
                {
                    WParagraph footerParagraph = (WParagraph)section.HeadersFooters.Footer.AddParagraph();

                    footerParagraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Center;

                    footerParagraph.AppendText("Page ");
                    footerParagraph.AppendField("Page", FieldType.FieldPage);

                    footerParagraph.AppendText(" of ");
                    footerParagraph.AppendField("NumPages", FieldType.FieldNumPages);
                }

                return CreateFileResult(document, "AddField.docx");
            }
        }

        public IActionResult UpdateField()
        {
            using (FileStream fileStream = new FileStream("Data\\Input.docx", FileMode.Open, FileAccess.Read))
            {
                WordDocument document = new WordDocument(fileStream, FormatType.Automatic);

                document.UpdateDocumentFields();

                return CreateFileResult(document, "UpdateField.docx");
            }
        }

        public IActionResult UnlinkField()
        {
            using (FileStream fileStream = new FileStream("Data\\InputTemplate.docx", FileMode.Open, FileAccess.Read))
            {
                WordDocument document = new WordDocument(fileStream, FormatType.Automatic);

                foreach (WSection section in document.Sections)
                {
                    foreach (WParagraph paragraph in section.Body.Paragraphs)
                    {
                        for (int i = 0; i < paragraph.Items.Count; i++)
                        {
                            if (paragraph.Items[i] is WField field)
                            {
                                if (field.FieldType == FieldType.FieldDate)
                                {
                                    field.Unlink();
                                }
                            }
                        }
                    }
                }

                return CreateFileResult(document, "UnlinkField.docx");
            }
        }

        private FileStreamResult CreateFileResult(WordDocument document, string fileName)
        {
            MemoryStream outputStream = new MemoryStream();
            document.Save(outputStream, FormatType.Docx);
            outputStream.Position = 0;
            return File(outputStream, "application/docx", fileName);
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
