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
                using (WordDocument document = new WordDocument(inputStream, FormatType.Automatic))
                {
                    //Hooks the font substitution event.
                    document.FontSettings.SubstituteFont += FontSettings_SubstituteFont;
                    //Creates a new instance of DocIORenderer class.
                    using (DocIORenderer render = new DocIORenderer())
                    {
                        //Converts Word document into PDF document.
                        using (PdfDocument pdfDocument = render.ConvertToPDF(document))
                        {
                            //Unhooks the font substitution event.
                            document.FontSettings.SubstituteFont -= FontSettings_SubstituteFont;
                            //Saves the PDF document to MemoryStream.
                            pdfDocument.Save(stream);
                            stream.Position = 0;
                        }
                    }
                }
            }
            //Download Word document in the browser.
            return File(stream, "application/pdf", "WordtoPDF.pdf");
            }
            catch(Exception ex)
            {
               ViewBag.Message = ex.Message;
            }
            return View();
        }
        private void FontSettings_SubstituteFont(object sender, SubstituteFontEventArgs args)
        {
            //Sets the alternate font when a specified font is not installed in the production environment.
            if (args.OriginalFontName == "Calibri" && args.FontStyle == FontStyle.Regular)
                args.AlternateFontStream = new FileStream(Path.GetFullPath(@"Data/calibri.ttf"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            else if (args.OriginalFontName == "Calibri" && args.FontStyle == FontStyle.Bold)
                args.AlternateFontStream = new FileStream(Path.GetFullPath(@"Data/calibrib.ttf"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            else if (args.OriginalFontName == "Segoe UI Light" && args.FontStyle == FontStyle.Regular)
                args.AlternateFontStream = new FileStream(Path.GetFullPath(@"Data/segoeuil.ttf"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            else if (args.OriginalFontName == "Segoe UI" && args.FontStyle == FontStyle.Regular)
                args.AlternateFontStream = new FileStream(Path.GetFullPath(@"Data/segoeui.ttf"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            else if (args.OriginalFontName == "Segoe UI" && args.FontStyle == FontStyle.Bold)
                args.AlternateFontStream = new FileStream(Path.GetFullPath(@"Data/segoeuib.ttf"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            else if (args.OriginalFontName == "Wingdings" && args.FontStyle == FontStyle.Regular)
                args.AlternateFontStream = new FileStream(Path.GetFullPath(@"Data/wingding.ttf"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
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