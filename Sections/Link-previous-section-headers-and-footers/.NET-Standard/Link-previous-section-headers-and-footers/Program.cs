using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;

namespace Link_previous_section_headers_and_footers
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
                //Inserts the first section header.
                section.HeadersFooters.Header.AddParagraph().AppendText("[ First Section Header ]");
                //Inserts the first section footer.
                section.HeadersFooters.Footer.AddParagraph().AppendText("[ First Section Footer ]");
                //Adds a paragraph to the section.
                IWParagraph paragraph = section.AddParagraph();
                string paraText = "AdventureWorks Cycles, the fictitious company on which the AdventureWorks sample databases are based, is a large, multinational manufacturing company.";
                //Appends some text to the first page in document.
                paragraph.AppendText("\r\r[ First Page ] \r\r" + paraText);
                //Adds the second section to the document.
                section = document.AddSection();
                //Inserts the second section header.
                section.HeadersFooters.Header.AddParagraph().AppendText("[ Second Section Header ]");
                //Inserts the second section footer.
                section.HeadersFooters.Footer.AddParagraph().AppendText("[ Second Section Footer ]");
                //Sets LinkToPrevious as true for retrieve the header and footer from previous section.
                section.HeadersFooters.LinkToPrevious = true;
                //Appends some text to the second page in document.
                paragraph = section.AddParagraph();
                paragraph.AppendText("\r\r[ Second Page ] \r\r" + paraText);
                //Adds the third section to the document.
                section = document.AddSection();
                //Inserts the third section header.
                section.HeadersFooters.Header.AddParagraph().AppendText("[ Third Section Header ]");
                //Inserts the third section footer.
                section.HeadersFooters.Footer.AddParagraph().AppendText("[ Third Section Footer ]");
                //Appends some text to the third page in document.
                paragraph = section.AddParagraph();
                paragraph.AppendText("\r\r[ Third Page ] \r\r" + paraText);
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
