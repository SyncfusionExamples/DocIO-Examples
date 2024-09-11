using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;

namespace Document_with_different_page_settings
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
                //Adds a paragraph to created section.
                IWParagraph paragraph = section.AddParagraph();
                string paraText = "AdventureWorks Cycles, the fictitious company on which the AdventureWorks sample databases are based, is a large, multinational manufacturing company.";
                //Appends the text to the created paragraph.
                paragraph.AppendText(paraText);
                //Sets the page size.
                section.PageSetup.PageSize = PageSize.A4;
                //Sets the page orientation as portrait.
                section.PageSetup.Orientation = PageOrientation.Portrait;
                //Adds the new section to the document.
                section = document.AddSection();
                //Sets the section break.
                section.BreakCode = SectionBreakCode.NewPage;
                paragraph = section.AddParagraph();
                //Sets the page size.
                section.PageSetup.PageSize = PageSize.A4;
                //Sets the page orientation as land scape.
                section.PageSetup.Orientation = PageOrientation.Landscape;
                //Appends the text to the paragraph.
                paragraph.AppendText(paraText);
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
