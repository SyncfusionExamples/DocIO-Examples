using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;

namespace Add_page_number_in_footer
{
    class Program
    {
        static void Main(string[] args)
        {
            //Creates a new Word document.
            using (WordDocument document = new WordDocument())
            {
                //Adds the section into Word document.
                IWSection section = document.AddSection();
                section.PageSetup.PageStartingNumber = 1;
                section.PageSetup.RestartPageNumbering = true;
                section.PageSetup.PageNumberStyle = PageNumberStyle.Arabic;
                //Specifies the value to header distance.
                section.PageSetup.HeaderDistance = 50;
                //Specifies the value to footer distance.
                section.PageSetup.FooterDistance = 50;
                //Adds a footer paragraph text to the document.
                IWParagraph paragraph = section.HeadersFooters.Footer.AddParagraph();
                paragraph.ParagraphFormat.Tabs.AddTab(523f, TabJustification.Right, TabLeader.NoLeader);
                // Adds text for the footer paragraph.
                paragraph.AppendText("Copyright Northwind Inc. 2001 - 2015\t");
                //Adds the text.
                paragraph.AppendText(" Page ");
                //Adds page number field to the document.
                paragraph.AppendField("CurrentPageNumber", FieldType.FieldPage);
                //Adds the text.
                paragraph.AppendText(" of ");
                //Adds number of page field to the document.
                paragraph.AppendField("TotalNumberOfPages", FieldType.FieldNumPages);
                //Adds a paragraph to the section.
                paragraph = section.AddParagraph();
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
