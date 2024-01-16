using Microsoft.AspNetCore.Mvc;
using Show_track_changes_mark_up.Models;
using System.Diagnostics;
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.DocIORenderer;
using Syncfusion.Pdf;
using System.Reflection.Metadata;

namespace Show_track_changes_mark_up.Controllers
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
        public ActionResult ConvertWordtoPDF(string TrackChangesOptions, string button)
        {
            //Open the file as Stream.
            using (FileStream docStream = new FileStream(Path.GetFullPath("Data/Template.docx"), FileMode.Open, FileAccess.Read))
            {
                //Loads file stream into Word document.
                using (WordDocument wordDocument = new WordDocument(docStream, FormatType.Automatic))
                {
                    #region Track changes option
                    //Sets revision types to preserve simple markup track changes in Word to PDF conversion.
                    if (TrackChangesOptions == "0")
                    {
                        wordDocument.RevisionOptions.ShowMarkup = RevisionType.Deletions | RevisionType.Formatting | RevisionType.Insertions | RevisionType.MoveFrom | RevisionType.MoveTo | RevisionType.StyleDefinitionChange;
                        wordDocument.RevisionOptions.ShowInBalloons = RevisionType.None;
                    }
                    //Sets revision types to preserve all markup track changes in Word to PDF conversion.
                    else if (TrackChangesOptions == "1")
                    {
                        wordDocument.RevisionOptions.ShowMarkup = RevisionType.Deletions | RevisionType.Formatting | RevisionType.Insertions | RevisionType.MoveFrom | RevisionType.MoveTo | RevisionType.StyleDefinitionChange;
                    }
                    //Sets none revision type to preserve no markup track changes in Word to PDF conversion.
                    else if (TrackChangesOptions == "2")
                    {
                        wordDocument.RevisionOptions.ShowMarkup = RevisionType.None;
                    }
                    #endregion
                    //Instantiation of DocIORenderer for Word to PDF conversion.
                    using (DocIORenderer render = new DocIORenderer())
                    {
                        //Converts Word document into PDF document.
                        PdfDocument pdfDocument = render.ConvertToPDF(wordDocument);

                        //Saves the PDF document to MemoryStream.
                        MemoryStream stream = new MemoryStream();
                        pdfDocument.Save(stream);
                        stream.Position = 0;

                        //Download PDF document in the browser.
                        return File(stream, "application/pdf", "Sample.pdf");
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
