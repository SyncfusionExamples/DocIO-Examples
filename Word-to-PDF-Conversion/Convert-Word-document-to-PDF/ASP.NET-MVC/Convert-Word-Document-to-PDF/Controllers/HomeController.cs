using Syncfusion.DocIO.DLS;
using Syncfusion.DocIO;
using Syncfusion.DocToPDFConverter;
using Syncfusion.Pdf;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace Convert_Word_Document_to_PDF.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            return View();
        }
        public ActionResult ConvertWordtoPDF()
        {
            //Open the file as Stream
            using (FileStream docStream = new FileStream(Server.MapPath("~/App_Data/Template.docx"), FileMode.Open, FileAccess.Read))
            {
                //Loads file stream into Word document
                using (WordDocument wordDocument = new WordDocument(docStream, FormatType.Automatic))
                {
                    //Instantiation of DocToPDFConverter for Word to PDF conversion
                    using (DocToPDFConverter converter = new DocToPDFConverter())
                    {
                        //Converts Word document into PDF document
                        using (PdfDocument pdfDocument = converter.ConvertToPDF(wordDocument))
                        {
                            //Saves the PDF document to MemoryStream.
                            MemoryStream stream = new MemoryStream();
                            pdfDocument.Save(stream);
                            stream.Position = 0;

                            //Download PDF document in the browser.
                            return File(stream, "application/pdf", "Sample.pdf");
                        }
                    };
                }
            }
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