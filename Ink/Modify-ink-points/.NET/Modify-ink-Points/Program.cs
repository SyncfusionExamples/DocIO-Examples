using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.Drawing;
using Syncfusion.Office;

namespace Modify_ink_points
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
                    // Access the ink and customize its trace points.
                    WInk ink = section.Paragraphs[0].ChildEntities[0] as WInk;
                    // Gets the ink trace from the ink object.
                    IOfficeInkTrace inkTrace = ink.Traces[0];
                    // Close the ink stroke by setting the last point to be the same as the first point
                    inkTrace.Points[inkTrace.Points.Length - 1] = new PointF(inkTrace.Points[0].X, inkTrace.Points[0].Y);
                    // Creates a file stream to save the modified document.
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