using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.Drawing;
using System.IO;

namespace Apply_page_borders
{
    class Program
    {
        static void Main(string[] args)
        {
            //Create a new Word document.
            using (WordDocument document = new WordDocument())
            {
                //Add a section to the Word document.
                IWSection section = document.AddSection();
                //Set the borders style.
                section.PageSetup.Borders.BorderType = BorderStyle.Single;
                //Set the color of the borders.
                section.PageSetup.Borders.Color = Color.Blue;
                //Set the linewidth of the borders.
                section.PageSetup.Borders.LineWidth = 0.75f;
                //Set the page border margins.
                section.PageSetup.Borders.Top.Space = 5f;
                section.PageSetup.Borders.Bottom.Space = 5f;
                section.PageSetup.Borders.Right.Space = 5f;
                section.PageSetup.Borders.Left.Space = 5f;
                //Add a paragraph to a section.
                IWParagraph paragraph = section.AddParagraph();
                paragraph.AppendText("AdventureWorks Cycles, the fictitious company on which the AdventureWorks sample databases are based, is a large, multinational manufacturing company.");
                //Create a file stream.
                using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"../../../Sample.docx"), FileMode.Create, FileAccess.ReadWrite))
                {
                    //Save the Word document to file the stream.
                    document.Save(outputFileStream, FormatType.Docx);
                }
            }
        }
    }
}

