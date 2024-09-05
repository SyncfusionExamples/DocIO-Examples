using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;

namespace Simple_headers_and_footers
{
    class Program
    {
        static void Main(string[] args)
        {
            //Creates a new Word document.
            using (WordDocument document = new WordDocument())
            {
                //Adds the first section to the document
                IWSection section = document.AddSection();
                //Adds a paragraph to the section
                IWParagraph paragraph = section.AddParagraph();
                string paraText = "AdventureWorks Cycles, the fictitious company on which the AdventureWorks sample databases are based, is a large, multinational manufacturing company.";
                //Appends some text to the first page in document
                paragraph.AppendText("\r\r[ First Page ] \r\r" + paraText);
                paragraph.ParagraphFormat.PageBreakAfter = true;
                //Appends some text to the second page in document
                paragraph = section.AddParagraph();
                paragraph.AppendText("\r\r[ Second Page ] \r\r" + paraText);
                paragraph.ParagraphFormat.PageBreakAfter = true;
                //Appends some text to the third page in document
                paragraph = section.AddParagraph();
                paragraph.AppendText("\r\r[ Third Page ] \r\r" + paraText);
                //Inserts the default page header
                paragraph = section.HeadersFooters.OddHeader.AddParagraph();
                paragraph.AppendText("[ Default Page Header ]");
                //Inserts the default Page footer
                paragraph = section.HeadersFooters.OddFooter.AddParagraph();
                paragraph.AppendText("[ Default Page Footer ]");
                //Creates file stream.
                using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"Output/Result.docx"), FileMode.Create, FileAccess.ReadWrite))
                {
                    //Saves the Word document to file stream.
                    document.Save(outputFileStream, FormatType.Docx);
                }
            }
        }
    }
}
