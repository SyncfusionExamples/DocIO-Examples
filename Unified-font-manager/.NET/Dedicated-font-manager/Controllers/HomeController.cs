using Dedicated_font_manager.Models;
using Microsoft.AspNetCore.Mvc;
using Syncfusion.DocIO.DLS;
using Syncfusion.DocIORenderer;
using Syncfusion.Pdf;
using Syncfusion.Presentation;
using Syncfusion.PresentationRenderer;
using Syncfusion.XlsIO;
using Syncfusion.XlsIORenderer;
using System.Diagnostics;

namespace Dedicated_font_manager.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;

        public HomeController(ILogger<HomeController> logger)
        {
            _logger = logger;
        }
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
                string fileName = Path.GetFileNameWithoutExtension(Request.Form.Files[0].FileName);

                // Compares extension with supported extensions.
                if (extension == ".doc" || extension == ".docx" || extension == ".dot" || extension == ".dotx" || extension == ".dotm" || extension == ".docm"
                   || extension == ".xml" || extension == ".rtf")
                {
                    try
                    {
                        MemoryStream outputStream = new MemoryStream();
                        //Open the Word document file stream.
                        using (MemoryStream inputStream = new MemoryStream())
                        {
                            Request.Form.Files[0].CopyTo(inputStream);
                            inputStream.Position = 0;
                            //Loads an existing Word document.
                            using (WordDocument wordDocument = new WordDocument())
                            {
                                wordDocument.Open(inputStream, Syncfusion.DocIO.FormatType.Automatic);
                                //Creates an instance of DocIORenderer.
                                using (DocIORenderer renderer = new DocIORenderer())
                                {
                                    //Converts Word document into PDF document.
                                    using (PdfDocument pdfDocument = renderer.ConvertToPDF(wordDocument))
                                    {
                                        pdfDocument.Save(outputStream);
                                    }
                                }
                            }
                        }
                        outputStream.Position = 0;
                        return File(outputStream, "application/pdf", fileName + ".pdf");
                    }
                    catch (Exception ex)
                    {
                        ViewBag.Message = string.Format(ex.Message);
                        //ViewBag.Message = string.Format("The input document could not be processed completely, Could you please email the document to support@syncfusion.com for troubleshooting.");
                    }
                }
                else if (extension == ".xlsx")
                {
                    try
                    {
                        using (MemoryStream inputStream = new MemoryStream())
                        {
                            Request.Form.Files[0].CopyTo(inputStream);
                            inputStream.Position = 0;
                            //Loads an existing Excel document.
                            using (ExcelEngine excelEngine = new ExcelEngine())
                            {
                                IApplication application = excelEngine.Excel;
                                application.DefaultVersion = ExcelVersion.Xlsx;
                                //Open the Excel document file stream.
                                IWorkbook workbook = application.Workbooks.Open(inputStream);
                                //Initialize XlsIO renderer.
                                XlsIORenderer renderer = new XlsIORenderer();
                                //Convert Excel document into PDF document 
                                PdfDocument pdfDocument = renderer.ConvertToPDF(workbook);
                                //Create the MemoryStream to save the converted PDF.      
                                MemoryStream pdfStream = new MemoryStream();
                                //Save the converted PDF document to MemoryStream.
                                pdfDocument.Save(pdfStream);
                                pdfStream.Position = 0;

                                //Download PDF document in the browser.
                                return File(pdfStream, "application/pdf", "Sample.pdf");
                            }
                        }

                    }
                    catch (Exception ex)
                    {
                        ViewBag.Message = string.Format(ex.Message);
                    }
                    
                }
                else if (extension == ".pptx")
                {
                    try 
                    {
                        using (MemoryStream inputStream = new MemoryStream())
                        {
                            Request.Form.Files[0].CopyTo(inputStream);
                            inputStream.Position = 0;
                            //Open the existing PowerPoint presentation with loaded stream.
                            using (IPresentation pptxDoc = Presentation.Open(inputStream))
                            {
                                //Convert the PowerPoint presentation to PDF document.
                                using (PdfDocument pdfDocument = PresentationToPdfConverter.Convert(pptxDoc))
                                {
                                    //Create the MemoryStream to save the converted PDF.      
                                    MemoryStream pdfStream = new MemoryStream();
                                    //Save the converted PDF document to MemoryStream.
                                    pdfDocument.Save(pdfStream);
                                    pdfStream.Position = 0;
                                    //Download PDF document in the browser.
                                    return File(pdfStream, "application/pdf", "Sample.pdf");
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        ViewBag.Message = string.Format(ex.Message);
                    }
                    
                }
                else
                {
                    ViewBag.Message = string.Format("Please choose Word, Excel or PowerPoint document to convert to PDF");
                }
            }
            else
            {
                ViewBag.Message = string.Format("Browse a Word,Excel or PowerPoint document and then click the button to convert as a PDF document");
            }
            return View("Index");
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
