using Convert_Word_Document_to_PDF.Models;
using Microsoft.AspNetCore.Mvc;
using Syncfusion.DocIO.DLS;
using Syncfusion.DocIO;
using Syncfusion.DocIORenderer;
using Syncfusion.Drawing;
using Syncfusion.Pdf;
using System.Diagnostics;

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
        public IActionResult ConvertWordtoPDF()
        {
            try
            {
                using (FileStream fileStreamPath = new FileStream(Path.GetFullPath("wwwroot/Data/Input.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    //Loads the template document.
                    using (WordDocument document = new WordDocument(fileStreamPath, FormatType.Docx))
                    {
                        //Hooks the font substitution event.
                        document.FontSettings.SubstituteFont += FontSettings_SubstituteFont;
                        using (DocIORenderer render = new DocIORenderer())
                        {
                            // Converts Word document into PDF document. 
                            using (PdfDocument pdf = render.ConvertToPDF(document))
                            {
                                MemoryStream memoryStream = new MemoryStream();
                                //Saves the PDF file.
                                pdf.Save(memoryStream);
                                //Unhooks the font substitution event after converting to PDF.
                                document.FontSettings.SubstituteFont -= FontSettings_SubstituteFont;
                                memoryStream.Position = 0;
                                //Download PDF document in the browser
                                return File(memoryStream, "application/pdf", "Sample.pdf");
                            }                               
                        }                           
                    }                      
                }                 
            }
            catch (Exception ex)
            {
                ViewBag.Message = ex.ToString();
            }
            return View("Index");
        }
        private static void FontSettings_SubstituteFont(object sender, SubstituteFontEventArgs args)
        {
            if (args.OrignalFontName == "Calibri" && args.FontStyle == FontStyle.Regular)
            {
                args.AlternateFontStream = new FileStream(Path.GetFullPath("wwwroot/Fonts/calibri.ttf"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
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