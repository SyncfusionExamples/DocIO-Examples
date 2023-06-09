using Convert_Word_Document_to_Image.Models;
using Microsoft.AspNetCore.Mvc;
using System.Diagnostics;
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.DocIORenderer;
using System;
using System.IO;
using SkiaSharp;

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
        public IActionResult ConvertWordtoImage()
        {
            //Open the file as Stream
            using (FileStream docStream = new FileStream(Path.GetFullPath("Data/Input.docx"), FileMode.Open, FileAccess.Read))
            {
                //Loads file stream into Word document
                using (WordDocument wordDocument = new WordDocument(docStream, FormatType.Automatic))
                {
                    //Instantiation of DocIORenderer
                    using (DocIORenderer render = new DocIORenderer())
                    {
                        //Convert the first page of the Word document into an image.
                        Stream imageStream = wordDocument.RenderAsImages(0, ExportImageFormat.Jpeg);
                        //Reset the stream position.
                        imageStream.Position = 0;
                        //Save the image file.
                       return File(imageStream, "application/jpeg", "wordtoimage.jpeg");
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