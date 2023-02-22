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
            Syncfusion.Licensing.SyncfusionLicenseProvider.RegisterLicense("Mgo+DSMBMAY9C3t2VVhkQlFadV5JXGFWfVJpTGpQdk5xdV9DaVZUTWY/P1ZhSXxRd0djXn5ZcXVQRWVfVEA=");
            //Create a new Word document.
            using (WordDocument document = new WordDocument())
            {
                //Add a section to the document.
                IWSection section = document.AddSection();
                //Set the borders style.
                section.PageSetup.Borders.BorderType = BorderStyle.Single;
                //Set the color of the borders.
                section.PageSetup.Borders.Color = Color.Black;
                //Set the linewidth of the borders.
                section.PageSetup.Borders.LineWidth = 0.5f;
                //Set whether the borders should be drawn with shadows.
                section.PageSetup.Borders.Shadow = false;
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
