using Convert_Word_Document_to_Image.Models;
using Microsoft.AspNetCore.Mvc;
using System.Diagnostics;
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.DocIORenderer;
using Microsoft.AspNetCore.Hosting;
using IHostingEnvironment = Microsoft.AspNetCore.Hosting.IHostingEnvironment;

namespace Convert_Word_Document_to_Image.Controllers
{
    public class HomeController : Controller
    {
        private IHostingEnvironment _env;
        public HomeController(IHostingEnvironment env)
        {
            _env = env;
        }

        public IActionResult Index()
        {
            return View();
        }
        /// <summary>
        /// Convert Word document to Image
        /// </summary>
        /// <param name="button"></param>
        /// <returns></returns>
        public IActionResult WordToImage(string button)
        {

            if (button == null)
                return View("Index");

            if (Request.Form.Files != null)
            {
                if (Request.Form.Files.Count == 0)
                {
                    ViewBag.Message = string.Format("Browse a Word document and then click the button to convert as a Image");
                    return View("Index");
                }
                // Gets the extension from file.
                string extension = Path.GetExtension(Request.Form.Files[0].FileName).ToLower();
                // Compares extension with supported extensions.
                if (extension == ".docx")
                {
                    MemoryStream stream = new MemoryStream();
                    Request.Form.Files[0].CopyTo(stream);
                    try
                    {
                        //Open using Syncfusion
                        using (WordDocument document = new WordDocument(stream, FormatType.Docx))
                        {
                            stream.Dispose();
                            // Creates a new instance of DocIORenderer class.
                            using (DocIORenderer render = new DocIORenderer())
                            {

                                //Convert the first page of the Word document into an image.
                                Stream imageStream = document.RenderAsImages(0, ExportImageFormat.Jpeg);
                                //Reset the stream position.
                                imageStream.Position = 0;
                                //Save the image file.
                                return File(imageStream, "application/jpeg", "sample.jpeg");
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        ViewBag.Message = ex.ToString();
                    }
                }
                else
                {
                    ViewBag.Message = string.Format("Please choose Word format document to convert to Image");
                }
            }
            else
            {
                ViewBag.Message = string.Format("Browse a Word document and then click the button to convert as a Image document");
            }
            return View("Index");
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