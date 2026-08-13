using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.Drawing;
using Syncfusion.Office;

namespace Modify_ink_color
{
    class Program
    {
        static void Main(string[] args)
        {
            using (FileStream fileStreamPath = new FileStream(Path.GetFullPath(@"Data/Template.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                //Creates a new Word document.
                using (WordDocument document = new WordDocument(fileStreamPath, FormatType.Docx))
                {
                    // Gets the first section of the document.
                    WSection section = document.Sections[0];
                    // Access the ink and customize its color.
                    WInk ink = section.Paragraphs[0].ChildEntities[0] as WInk;
                    // Gets the ink trace from the ink object.
                    IOfficeInkTrace inkTrace = ink.Traces[0];
                    // Modify the brush color to Color.Red.
                    inkTrace.Brush.Color = Color.Red;
                    using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"../../../Output/Result.docx"), FileMode.Create, FileAccess.ReadWrite))
                    {
                        //Saves the Word document to file stream.
                        document.Save(outputFileStream, FormatType.Docx);
                    }
                }
            }
        }
    }
}