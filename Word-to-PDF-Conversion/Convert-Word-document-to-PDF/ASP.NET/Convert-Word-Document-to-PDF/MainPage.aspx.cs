using System;
using System.Web;
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;
using Syncfusion.DocToPDFConverter;
using Syncfusion.Pdf;
using System.Web.Configuration;

namespace Convert_Word_Document_to_PDF
{
    public partial class WebForm1 : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {

        }
        protected void OnButtonClicked(object sender, EventArgs e)
        {   
            //Open an existing Word document.        
            string filePath = Server.MapPath("~/App_Data/Template.docx");

            //Loads file into Word document
            using (WordDocument document = new WordDocument(filePath))
            {
                //Instantiation of DocToPDFConverter for Word to PDF conversion
                using (DocToPDFConverter converter = new DocToPDFConverter())
                {
                    //Converts Word document into PDF document
                    using (PdfDocument pdfDocument = converter.ConvertToPDF(document))
                    {
                        //Saves the PDF document to MemoryStream.
                        MemoryStream stream = new MemoryStream();
                        pdfDocument.Save("sample.pdf", HttpContext.Current.Response, HttpReadType.Save);                       
                    }                   
                }
            }
        }
    }
}