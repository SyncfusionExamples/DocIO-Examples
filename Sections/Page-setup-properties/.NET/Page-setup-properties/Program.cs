using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;

namespace Page_setup_properties
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
                //Sets page setup options.
                section.PageSetup.Orientation = PageOrientation.Landscape;
                section.PageSetup.Margins.All = 72;
                section.PageSetup.Borders.LineWidth = 2;
                //Sets the PrinterPaperTray value for FirstPageTray in page setup options.
                section.PageSetup.FirstPageTray = PrinterPaperTray.EnvelopeFeed;
                //Sets the PrinterPaperTray value for OtherPagesTray in page setup options.
                section.PageSetup.OtherPagesTray = PrinterPaperTray.MiddleBin;
                //Adds a paragraph to created section.
                IWParagraph paragraph = section.AddParagraph();
                //Appends the text to the created paragraph.
                paragraph.AppendText("AdventureWorks Cycles, the fictitious company on which the AdventureWorks sample databases are based, is a large, multinational manufacturing company.");
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
