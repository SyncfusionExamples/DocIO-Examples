using Convert_Word_Document_to_Image.Models;
using Microsoft.AspNetCore.Mvc;
using System.Diagnostics;
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.DocIORenderer;
using System.Reflection.Metadata;
using Syncfusion.Drawing;

namespace Convert_Word_Document_to_Image.Controllers
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

        public ActionResult ConvertWordtoImage()
        {
            //Open the file as Stream
            using (FileStream docStream = new FileStream(Path.GetFullPath("wwwroot/Data/Input.docx"), FileMode.Open, FileAccess.Read))
            {
                //Loads file stream into Word document
                using (WordDocument wordDocument = new WordDocument(docStream, FormatType.Docx))
                {
                    //Hooks the font substitution event.
                    wordDocument.FontSettings.SubstituteFont += FontSettings_SubstituteFont;
                    //Instantiation of DocIORenderer
                    using (DocIORenderer render = new DocIORenderer())
                    {
                        //Convert the first page of the Word document into an image.
                        Stream imageStream = wordDocument.RenderAsImages(0, ExportImageFormat.Jpeg);
                        //Unhooks the font substitution event after converting to image.
                        wordDocument.FontSettings.SubstituteFont -= FontSettings_SubstituteFont;
                        //Reset the stream position.
                        imageStream.Position = 0;
                        //Save the image file.
                        return File(imageStream, "application/jpeg", "WordToImage.Jpeg");
                    }
                }
            }
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