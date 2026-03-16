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
        public IActionResult OfficeToPDF(string button)
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

                try
                {
                    PdfDocument pdfDocument = null;
                    // Switch on file extension to determine conversion method
                    switch (extension.ToLower())
                    {
                        // Word document formats
                        case ".doc":
                        case ".docx":
                        case ".dot":
                        case ".dotx":
                        case ".dotm":
                        case ".docm":
                        case ".xml":
                        case ".rtf":
                            using (MemoryStream inputStream = new MemoryStream())
                            {
                                // Copy uploaded file to memory stream
                                Request.Form.Files[0].CopyTo(inputStream);
                                inputStream.Position = 0;
                                // Open and load the Word document
                                using (WordDocument wordDocument = new WordDocument())
                                {
                                    // Convert Word document to PDF format
                                    wordDocument.Open(inputStream, Syncfusion.DocIO.FormatType.Automatic);
                                    using (DocIORenderer renderer = new DocIORenderer())
                                    {
                                        pdfDocument = renderer.ConvertToPDF(wordDocument);
                                    }
                                }
                            }
                            break;
                        // Excel format
                        case ".xlsx":
                        case ".xls":
                        case ".xltx":
                        case ".xlsm":
                        case ".csv":
                        case ".xlsb":
                        case ".xltm":
                            using (MemoryStream inputStream = new MemoryStream())
                            {
                                // Copy uploaded file to memory stream
                                Request.Form.Files[0].CopyTo(inputStream);
                                inputStream.Position = 0;
                                // Create Excel engine and load workbook
                                using (ExcelEngine excelEngine = new ExcelEngine())
                                {
                                    IApplication application = excelEngine.Excel;
                                    application.DefaultVersion = ExcelVersion.Xlsx;
                                    IWorkbook workbook = application.Workbooks.Open(inputStream);
                                    // Convert Excel workbook to PDF format
                                    XlsIORenderer renderer = new XlsIORenderer();
                                    pdfDocument = renderer.ConvertToPDF(workbook);
                                }
                            }
                            break;
                        // PowerPoint format
                        case ".pptx":
                            using (MemoryStream inputStream = new MemoryStream())
                            {
                                // Copy uploaded file to memory stream
                                Request.Form.Files[0].CopyTo(inputStream);
                                inputStream.Position = 0;
                                // Open PowerPoint presentation and convert to PDF
                                using (IPresentation pptxDoc = Presentation.Open(inputStream))
                                {
                                    pdfDocument = PresentationToPdfConverter.Convert(pptxDoc);
                                }
                            }
                            break;
                        // Invalid file format
                        default:
                            ViewBag.Message = "Please choose Word, Excel or PowerPoint document to convert to PDF";
                            return null;
                    }
                    // Save converted PDF and return as downloadable file
                    if (pdfDocument != null)
                    {
                        using (pdfDocument)
                        {
                            // Create memory stream to hold the PDF data
                            MemoryStream pdfStream = new MemoryStream();
                            // Save the converted PDF to memory stream
                            pdfDocument.Save(pdfStream);
                            // Reset stream position to beginning for reading
                            pdfStream.Position = 0;
                            // Return PDF as downloadable file to browser
                            return File(pdfStream, "application/pdf", fileName + ".pdf");
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
