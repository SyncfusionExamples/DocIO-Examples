using Convert_Word_Document_to_Image.Models;
using Microsoft.AspNetCore.Mvc;
using System.Diagnostics;
using Syncfusion.DocIO.DLS;
using Syncfusion.DocIO;
using Syncfusion.DocIORenderer;
using Syncfusion.Drawing;
using Syncfusion.Pdf;

namespace Convert_Word_Document_to_Image.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;

        public HomeController(ILogger<HomeController> logger)
        {
            _logger = logger;
        }

        public ActionResult ConvertWordToImage()
        {
            MemoryStream imageStream = new MemoryStream();
            try
            {
                //Open an existing Word document.
                using (FileStream inputStream = new FileStream(Path.GetFullPath("Data/Input.docx"), FileMode.Open, FileAccess.Read))
                {
                    //Create a new document.
                    using (WordDocument document = new WordDocument(inputStream, FormatType.Docx))
                    {
                        //Creates a new instance of DocIORenderer class.
                        using (DocIORenderer render = new DocIORenderer())
                        {
                            //Converts the first page of word document to image
                            imageStream = (MemoryStream)document.RenderAsImages(0, ExportImageFormat.Jpeg);
                        }
                    }
                }
                //Download Word document in the browser.
                return File(imageStream, "image/jpeg", "WordToimage_Page1.jpeg");
            }
            catch (Exception ex)
            {
                ViewBag.Message = ex.Message;
            }
            return View();
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