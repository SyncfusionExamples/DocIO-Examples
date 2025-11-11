using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.DocIORenderer;
using Syncfusion.Pdf;
using System.IO;

namespace Mail_merged_Word_into_PDF
{
    class Program
    {
        static void Main(string[] args)
        {
            //Opens the Word template document
            FileStream fileStreamPath = new FileStream(@"../../../Letter Formatting.docx", FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            using (WordDocument document = new WordDocument(fileStreamPath, FormatType.Docx))
            {
                string[] fieldNames = { "ContactName", "CompanyName", "Address", "City", "Country", "Phone" };
                string[] fieldValues = { "Nancy Davolio", "Syncfusion", "507 - 20th Ave. E.Apt. 2A", "Seattle, WA", "USA", "(206) 555-9857-x5467" };
                //Performs the mail merge
                document.MailMerge.Execute(fieldNames, fieldValues);
                //Create instance for DocIORenderer for Word to PDF conversion
                DocIORenderer render = new DocIORenderer();
                //Converts Word document to PDF.
                PdfDocument pdfDocument = render.ConvertToPDF(document);
                //Release the resources used by the Word document and DocIO Renderer objects.
                render.Dispose();
                document.Dispose();
                //Saves the PDF file.
                FileStream outputStream = new FileStream(@"Ouput.pdf", FileMode.CreateNew, FileAccess.Write);
                pdfDocument.Save(outputStream);
                //Closes the instance of PDF document object.
                pdfDocument.Close();
                //Dispose the instance of FileStream.
                outputStream.Dispose();
            }
        }
    }
}
