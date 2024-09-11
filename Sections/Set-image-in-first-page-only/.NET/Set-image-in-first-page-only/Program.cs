using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;

namespace Add_image
{
    class Program
    {
        static void Main(string[] args)
        {
            //Creates a new Word document.
            using (WordDocument document = new WordDocument())
            {
                //Adds new section to the document.
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
                //Appends some text to the second page in document.
                paragraph.AppendText("\r\r[ Second Page ] \r\r" + paraText);
                //Inserts the first page header.
                paragraph = section.HeadersFooters.FirstPageHeader.AddParagraph();
                //Adds image to the paragraph.
                FileStream imageStream = new FileStream(Path.GetFullPath(@"Data/Image.png"), FileMode.Open, FileAccess.ReadWrite);
                IWPicture picture = paragraph.AppendPicture(imageStream);
                //Sets the text wrapping style as Behind the text.
                picture.TextWrappingStyle = TextWrappingStyle.Behind;
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
