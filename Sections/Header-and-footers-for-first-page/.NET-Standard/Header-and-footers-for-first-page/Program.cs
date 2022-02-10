using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;

namespace Header_and_footers_for_first_page
{
    class Program
    {
        static void Main(string[] args)
        {
            //Creates a new Word document.
            using (WordDocument document = new WordDocument())
            {
                //Adds the first section to the document.
                IWSection section = document.AddSection();
                //Sets DifferentFirstPage as true for inserting header and footer text.
                section.PageSetup.DifferentFirstPage = true;
                //Adds a paragraph to the section.
                IWParagraph paragraph = section.AddParagraph();
                string paraText = "AdventureWorks Cycles, the fictitious company on which the AdventureWorks sample databases are based, is a large, multinational manufacturing company.";
                //Appends some text to the first page in document.
                paragraph.AppendText("\r\r[ First Page ] \r\r" + paraText);
                paragraph.ParagraphFormat.PageBreakAfter = true;
                //Appends some text to the second page in document.
                paragraph = section.AddParagraph();
                paragraph.AppendText("\r\r[ Second Page ] \r\r" + paraText);
                paragraph.ParagraphFormat.PageBreakAfter = true;
                //Appends some text to the third page in document.
                paragraph = section.AddParagraph();
                paragraph.AppendText("\r\r[ Third Page ] \r\r" + paraText);
                //Inserts the first page header.
                paragraph = section.HeadersFooters.FirstPageHeader.AddParagraph();
                paragraph.AppendText("[First Page Header ]");
                //Inserts the first page footer.
                paragraph = section.HeadersFooters.FirstPageFooter.AddParagraph();
                paragraph.AppendText("[ First Page Footer ]");
                //Inserts the default page header.
                paragraph = section.HeadersFooters.OddHeader.AddParagraph();
                paragraph.AppendText("[ Default Page Header ]");
                //Inserts the default page footer.
                paragraph = section.HeadersFooters.OddFooter.AddParagraph();
                paragraph.AppendText("[ Default Page Footer ]");
                //Creates file stream.
                using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"../../../Result.docx"), FileMode.Create, FileAccess.ReadWrite))
                {
                    //Saves the Word document to file stream.
                    document.Save(outputFileStream, FormatType.Docx);
                }
            }
        }
    }
}
