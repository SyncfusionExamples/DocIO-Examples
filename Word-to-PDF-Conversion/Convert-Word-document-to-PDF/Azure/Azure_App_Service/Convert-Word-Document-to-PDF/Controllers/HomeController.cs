using Convert_Word_Document_to_PDF.Models;
using Microsoft.AspNetCore.Mvc;
using System.Diagnostics;
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.DocIORenderer;
using Syncfusion.Pdf;
using Microsoft.AspNetCore.Hosting;
using System.IO;

namespace Convert_Word_Document_to_PDF.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;

        // [C# Code]
        private Microsoft.AspNetCore.Hosting.IHostingEnvironment _env;
        public HomeController(Microsoft.AspNetCore.Hosting.IHostingEnvironment env)
        {
            _env = env;
        }

        public IActionResult Index()
        {
            return View();
        }
        /// <summary>
        /// Convert Word document to PDF
        /// </summary>
        /// <param name="button"></param>
        /// <returns></returns>
        public IActionResult WordToPDF(string button)
        {
            if (button == null)
                return View("Index");

            if (Request.Form.Files != null)
            {
                if (Request.Form.Files.Count == 0)
                {
                    ViewBag.Message = string.Format("Browse a Word document and then click the button to convert as a PDF document");
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
                                // Converts Word document into PDF document
                                using (PdfDocument pdf = render.ConvertToPDF(document))
                                {                                                                     
                                    MemoryStream memoryStream = new MemoryStream();
                                    // Save the PDF document
                                    pdf.Save(memoryStream);
                                    memoryStream.Position = 0;
                                    return File(memoryStream, "application/pdf", "WordToPDF.pdf");
                                }                                                           
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
                    ViewBag.Message = string.Format("Please choose Word format document to convert to PDF");
                }
            }
            else
            {
                ViewBag.Message = string.Format("Browse a Word document and then click the button to convert as a PDF document");
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