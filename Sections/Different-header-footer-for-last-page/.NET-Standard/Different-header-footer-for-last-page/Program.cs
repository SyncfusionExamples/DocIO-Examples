using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;

namespace Different_header_footer_for_last_page
{
    class Program
    {
        static void Main(string[] args)
        {
            //Create a new Word document.
            using (WordDocument document = new WordDocument())
            {
                //Add the first section to the document.
                IWSection section = document.AddSection();
                //Add a paragraph to the section.
                IWParagraph paragraph = section.AddParagraph();
                string paraText = "AdventureWorks Cycles, the fictitious company on which the AdventureWorks sample databases are based, is a large, multinational manufacturing company.";
                //Append some text to the first page in document.
                paragraph.AppendText("\r[ First Page ] \r\r" + paraText);
                //Insert default header and footer to the section.
                paragraph = section.HeadersFooters.Header.AddParagraph();
                paragraph.AppendText("[ Header ]");
                paragraph = section.HeadersFooters.Footer.AddParagraph();
                paragraph.AppendText("[ Footer ]");
                //Add the second section to the document.
                section = document.AddSection();
                //Append some text to the second page in document.
                paragraph = section.AddParagraph();
                paragraph.AppendText("\r[ Second Page ] \r\r" + paraText);
                //Add the third section to the document.
                section = document.AddSection();
                //Append some text to the third page in document.
                paragraph = section.AddParagraph();
                paragraph.AppendText("\r[ Third Page ] \r\r" + paraText);
                //Insert default header and footer to the section.
                paragraph = section.HeadersFooters.Header.AddParagraph();
                paragraph.AppendText("[ Third Page Header ]");
                paragraph = section.HeadersFooters.Footer.AddParagraph();
                paragraph.AppendText("[ Third Page Footer ]");
                //Create a file stream.
                using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"../../../Sample.docx"), FileMode.Create, FileAccess.ReadWrite))
                {
                    //Save the Word document to the file stream.
                    document.Save(outputFileStream, FormatType.Docx);
                }
            }
        }
    }
}
