using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;

namespace Rotate_shapes
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
                WParagraph paragraph = section.AddParagraph() as WParagraph;
                Shape rectangle = paragraph.AppendShape(AutoShapeType.RoundedRectangle, 150, 100);
                //Sets position for shape.
                rectangle.VerticalPosition = 72;
                rectangle.HorizontalPosition = 72;
                //Sets 90 degree rotation.
                rectangle.Rotation = 90;
                //Sets horizontal flip.
                rectangle.FlipHorizontal = true;
                paragraph = section.AddParagraph() as WParagraph;
                paragraph = rectangle.TextBody.AddParagraph() as WParagraph;
                IWTextRange text = paragraph.AppendText("This text is in rounded rectangle shape");
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
