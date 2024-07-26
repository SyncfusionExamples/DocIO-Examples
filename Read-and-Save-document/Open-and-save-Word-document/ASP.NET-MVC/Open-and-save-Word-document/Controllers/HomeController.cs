using Syncfusion.DocIO.DLS;
using Syncfusion.DocIO;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace Open_and_save_Word_document.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            return View();
        }
        public ActionResult OpenAndSaveDocument()
        {
            //Open an existing Word document.
            using (WordDocument document = new WordDocument(Server.MapPath("~/App_Data/Input.docx")))
            {
                //Access the section in a Word document.
                IWSection section = document.Sections[0];
                //Add a new paragraph to the section.
                IWParagraph paragraph = section.AddParagraph();
                paragraph.ParagraphFormat.FirstLineIndent = 36;
                paragraph.BreakCharacterFormat.FontSize = 12f;
                IWTextRange text = paragraph.AppendText("In 2000, Adventure Works Cycles bought a small manufacturing plant, Importadores Neptuno, located in Mexico. Importadores Neptuno manufactures several critical subcomponents for the Adventure Works Cycles product line. These subcomponents are shipped to the Bothell location for final product assembly. In 2001, Importadores Neptuno, became the sole manufacturer and distributor of the touring bicycle product group.");
                text.CharacterFormat.FontSize = 12f;

                //Save the Word document and download as attachment.
                document.Save("Sample.docx", FormatType.Docx, HttpContext.ApplicationInstance.Response, HttpContentDisposition.Attachment);
            }

            return View();
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