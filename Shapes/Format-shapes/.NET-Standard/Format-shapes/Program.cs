using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.Drawing;
using System.IO;

namespace Format_shapes
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
                IWParagraph paragraph = section.AddParagraph() as WParagraph;
                Shape rectangle = paragraph.AppendShape(AutoShapeType.RoundedRectangle, 150, 100);
                rectangle.VerticalPosition = 72;
                rectangle.HorizontalPosition = 72;
                paragraph = section.AddParagraph() as WParagraph;
                paragraph = rectangle.TextBody.AddParagraph() as WParagraph;
                IWTextRange text = paragraph.AppendText("This text is in rounded rectangle shape");
                text.CharacterFormat.TextColor = Color.Green;
                text.CharacterFormat.Bold = true;
                //Applies fill color for shape.
                rectangle.FillFormat.Fill = true;
                rectangle.FillFormat.Color = Color.LightGray;
                //Applies wrap formats.
                rectangle.WrapFormat.TextWrappingStyle = TextWrappingStyle.Square;
                rectangle.WrapFormat.TextWrappingType = TextWrappingType.Right;
                //Sets horizontal and vertical origin.
                rectangle.HorizontalOrigin = HorizontalOrigin.Margin;
                rectangle.VerticalOrigin = VerticalOrigin.Page;
                //Sets line format.
                rectangle.LineFormat.DashStyle = LineDashing.Dot;
                rectangle.LineFormat.Color = Color.DarkGray;
                //Creates file stream.
                using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"../../../Result.docx"), FileMode.Create, FileAccess.ReadWrite))
                {
                    //Saves the Word document to file stream.
                    document.Save(outputFileStream, FormatType.Docx);
                }
            }
        }
    }
}
