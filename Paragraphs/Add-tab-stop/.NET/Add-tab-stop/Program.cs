using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.Drawing;
using System.IO;

namespace Add_tab_stop
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
                //Adds new paragraph to the section.
                IWParagraph paragraph = section.AddParagraph();
                //Adds tab stop at position 11.
                Tab firstTab = paragraph.ParagraphFormat.Tabs.AddTab(11, TabJustification.Left, TabLeader.Dotted);
                //Adds tab stop at position 62.
                paragraph.ParagraphFormat.Tabs.AddTab(62, TabJustification.Left, TabLeader.Single);
                paragraph.AppendText("This sample\t illustrates the use of tabs in the paragraph. Tabs\t can be inserted or removed from the paragraph.");
                //Removes tab stop from the collection.
                paragraph.ParagraphFormat.Tabs.RemoveByTabPosition(11);
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
