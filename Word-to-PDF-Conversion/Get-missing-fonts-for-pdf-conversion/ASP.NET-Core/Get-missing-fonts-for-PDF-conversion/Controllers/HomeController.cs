using Get_missing_fonts_for_PDF_conversion.Models;
using Microsoft.AspNetCore.Mvc;
using System.Diagnostics;
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.DocIORenderer;
using Syncfusion.Pdf;
using Microsoft.AspNetCore.Hosting;
using System.IO;

namespace Get_missing_fonts_for_PDF_conversion.Controllers
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
        // List to store names of fonts that are not installed
        static List<string> fonts = new List<string>();

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
                Stream stream = GetWordDocument();
                try
                {
                    //Open using Syncfusion
                    using (WordDocument document = new WordDocument(stream, FormatType.Docx))
                    {
                        stream.Dispose();
                        // Hook the font substitution event to detect missing fonts
                        document.FontSettings.SubstituteFont += FontSettings_SubstituteFont;
                        // Creates a new instance of DocIORenderer class.
                        using (DocIORenderer render = new DocIORenderer())
                        {
                            // Converts Word document into PDF document
                            using (PdfDocument pdf = render.ConvertToPDF(document))
                            {
                                // Print the fonts that are not available in machine, but used in Word document.
                                if (fonts.Count > 0)
                                {
                                    Console.WriteLine("Fonts not available in environment:");
                                    foreach (string font in fonts)
                                        Console.WriteLine(font);
                                }
                                else
                                {
                                    Console.WriteLine("Fonts used in Word document are available in environment.");
                                }

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
                ViewBag.Message = string.Format("Browse a Word document and then click the button to convert as a PDF document");
            }
            return View("Index");
        }

        // Event handler for font substitution event
        static void FontSettings_SubstituteFont(object sender, SubstituteFontEventArgs args)
        {
            // Add the original font name to the list if it's not already there
            if (!fonts.Contains(args.OriginalFontName))
                fonts.Add(args.OriginalFontName);
        }

        private Stream GetWordDocument()
        {
            if (Request.Form.Files != null && Request.Form.Files.Count != 0)
            {
                // Gets the extension from file.
                string extension = Path.GetExtension(Request.Form.Files[0].FileName).ToLower();

                // Compares extension with supported extensions.
                if (extension == ".doc" || extension == ".docx" || extension == ".dot" || extension == ".dotx" || extension == ".dotm" || extension == ".docm"
                   || extension == ".xml" || extension == ".rtf")
                {
                    MemoryStream stream = new MemoryStream();
                    Request.Form.Files[0].CopyTo(stream);
                    return stream;
                }
                else
                {
                    ViewBag.Message = string.Format("Please choose Word format document to convert to PDF");
                    return null;
                }
            }
            else
            {
                //Opens an existing document from stream through constructor of `WordDocument` class
                FileStream fileStreamPath = new FileStream(_env.WebRootPath + @"/Data/Input.docx", FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                return fileStreamPath;
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