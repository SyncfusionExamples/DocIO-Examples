using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using static System.Collections.Specialized.BitVector32;

namespace Send_Word_to_client_browser
{
    public partial class MainPage : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {

        }
        protected void OnButtonClicked(object sender, EventArgs e)
        {
            //Creating a new document.
            using (WordDocument document = new WordDocument())
            {
                 //Adds new section to the document.
                IWSection section = document.AddSection();
                //Adds new paragraph to the section.
                IWParagraph paragraph = section.AddParagraph();
                //Appends the text to the created paragraph.
                paragraph.AppendText("Adventure Works Cycles, the fictitious company on which the AdventureWorks sample databases are based, is a large, multinational manufacturing company.");

#region SaveAndDownloadDocument
                //Saves the Word document to disk in DOCX format.
                MemoryStream stream = new MemoryStream();
                document.Save(stream,FormatType.Docx);
                Response.Clear();
                Response.ContentType = "application/msword";
                Response.AddHeader("Content-Disposition", "attachment; filename=\"Sample.docx\"");
                stream.CopyTo(Response.OutputStream);
                Response.End();
#endregion
            }
        }
    }
}