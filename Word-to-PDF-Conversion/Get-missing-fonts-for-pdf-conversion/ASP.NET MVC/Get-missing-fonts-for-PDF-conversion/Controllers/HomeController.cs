using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.DocToPDFConverter;
using Syncfusion.Pdf;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace Create_Word_document.Controllers
{
    public class HomeController : Controller
    {
        // List to store names of fonts that are not installed
        static List<string> fonts = new List<string>();
        public ActionResult Index()
        {
            return View();
        }
        /// <summary>
        /// Convert Word document to PDF
        /// </summary>
        /// <param name="button"></param>
        /// <returns></returns>
        public ActionResult WordToPDF(string button)
        {
            if (button == null)
                return View("Index");

            if (Request.Files != null)
            {
                Stream stream = GetWordDocument();

                if (stream == null)
                {
                    return View("Index");
                }

                try
                {
                    //Open using Syncfusion
                    using (WordDocument document = new WordDocument(stream, FormatType.Docx))
                    {
                        stream.Dispose();
                        // Hook the font substitution event to detect missing fonts
                        document.FontSettings.SubstituteFont += FontSettings_SubstituteFont;
                        // Creates a new instance of DocIORenderer class.
                        using (DocToPDFConverter render = new DocToPDFConverter())
                        {
                            // Converts Word document into PDF document
                            using (PdfDocument pdf = render.ConvertToPDF(document))
                            {
                                // Print the fonts that are not available in machine, but used in Word document.

                                string dir = Server.MapPath("~/MissingFontDetails");
                                Directory.CreateDirectory(dir);
                                string filePath = Path.Combine(dir, "MissingFonts.txt");

                                if (fonts.Count > 0)
                                {
                                    var lines = new List<string>
                                    {
                                        "Fonts not available in environment:",
                                        "----------------------------------"
                                    };
                                    lines.AddRange(fonts.Distinct().OrderBy(f => f));

                                    System.IO.File.WriteAllLines(filePath, lines);
                                }
                                else
                                {
                                    System.IO.File.WriteAllText(filePath, "Fonts used in Word document are available in environment.");
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

        static void FontSettings_SubstituteFont(object sender, SubstituteFontEventArgs args)
        {
            // Add the original font name to the list if it's not already there
            if (!fonts.Contains(args.OriginalFontName))
                fonts.Add(args.OriginalFontName);
        }

        private Stream GetWordDocument()
        {
            if (Request.Files != null && Request.Files.Count > 0)
            {
                HttpPostedFileBase posted = Request.Files[0];

                // Ensure file is actually uploaded
                if (posted != null && posted.ContentLength > 0)
                {
                    string extension = Path.GetExtension(posted.FileName).ToLower();
                    if (extension == ".doc" || extension == ".docx" || extension == ".dot" || extension == ".dotx" ||
                        extension == ".dotm" || extension == ".docm" || extension == ".xml" || extension == ".rtf")
                    {
                        MemoryStream stream = new MemoryStream();
                        posted.InputStream.CopyTo(stream);
                        stream.Position = 0;
                        return stream;
                    }
                    else
                    {
                        ViewBag.Message = "Please choose a Word format document to convert to PDF";
                        return null;
                    }
                }
            }

            // Fallback to default document
            string defaultPath = Server.MapPath("~/Data/Input.docx");
            return new FileStream(defaultPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        }


        public ActionResult About()
        {
            ViewBag.Message = "Your application description page.";

            return View();
        }

        public ActionResult Contact()
        {
            ViewBag.Message = "Your contact page.";

            return View();
        }
    }
}