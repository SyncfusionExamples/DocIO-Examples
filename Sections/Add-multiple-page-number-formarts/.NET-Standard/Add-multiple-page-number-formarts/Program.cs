using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;

namespace Add_multiple_page_number_formarts
{
    class Program
    {
        static void Main(string[] args)
        {
            //Create a new Word document.
            using (WordDocument document = new WordDocument())
            {
                //Add a section to the document.
                IWSection section = document.AddSection();
                section.PageSetup.PageStartingNumber = 1;
                section.PageSetup.PageNumberStyle = PageNumberStyle.Arabic;
                //Add a footer paragraph to the document.
                IWParagraph paragraph = section.HeadersFooters.Footer.AddParagraph();
                paragraph.ParagraphFormat.Tabs.AddTab(523f, TabJustification.Left, TabLeader.NoLeader);
                //Add page number field to the document.
                paragraph.AppendText("Page ");
                paragraph.AppendField("Page", FieldType.FieldPage);
                //Add a paragraph to a section.
                paragraph = section.AddParagraph();
                //Append the text to the created paragraph.
                paragraph.AppendText("AdventureWorks Cycles, the fictitious company on which the AdventureWorks sample databases are based, is a large, multinational manufacturing company.");
                //Add a section to the document.
                section = document.AddSection();
                section.PageSetup.PageStartingNumber = 1;
                section.PageSetup.RestartPageNumbering = true;
                section.PageSetup.PageNumberStyle = PageNumberStyle.LetterUpper;
                //Add a paragraph to a section.
                paragraph = section.AddParagraph();
                paragraph.AppendText("AdventureWorks Cycles, the fictitious company on which the AdventureWorks sample databases are based, is a large, multinational manufacturing company.");
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
