using Syncfusion.DocIO.DLS;
using Syncfusion.DocIO;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace Open_and_save_Word_document
{
    public partial class MainPage : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {

        }
        protected void OnButtonClicked(object sender, EventArgs e)
        {
            string filePath = Server.MapPath("~/App_Data/Input.docx");
            //Open an existing Word document.
            using (WordDocument document = new WordDocument(filePath))
            {
                //Access the section in a Word document.
                IWSection section = document.Sections[0];
                //Add a new paragraph to the section.
                IWParagraph paragraph = section.AddParagraph();
                paragraph.ParagraphFormat.FirstLineIndent = 36;
                paragraph.BreakCharacterFormat.FontSize = 12f;
                IWTextRange text = paragraph.AppendText("In 2000, Adventure Works Cycles bought a small manufacturing plant, Importadores Neptuno, located in Mexico. Importadores Neptuno manufactures several critical subcomponents for the Adventure Works Cycles product line. These subcomponents are shipped to the Bothell location for final product assembly. In 2001, Importadores Neptuno, became the sole manufacturer and distributor of the touring bicycle product group.");
                text.CharacterFormat.FontSize = 12f;

                //Save the Word document to the disk in a DOCX format.
                document.Save("Sample.docx", FormatType.Docx, HttpContext.Current.Response, HttpContentDisposition.Attachment);
            }
        }
    }
}