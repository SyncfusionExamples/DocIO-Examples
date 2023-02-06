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
            //Create a new Word document.
            using (WordDocument document = new WordDocument())
            {
                //Add a new section to the document.
                IWSection section = document.AddSection();
                //Add a new paragraph to the section.
                IWParagraph paragraph = section.AddParagraph() as WParagraph;
                Shape rectangle = paragraph.AppendShape(AutoShapeType.RoundedRectangle, 150, 100);
                rectangle.VerticalPosition = 72;
                rectangle.HorizontalPosition = 72;
                paragraph = section.AddParagraph() as WParagraph;
                paragraph = rectangle.TextBody.AddParagraph() as WParagraph;
                IWTextRange text = paragraph.AppendText("This text is in rounded rectangle shape");
                text.CharacterFormat.TextColor = Color.Green;
                text.CharacterFormat.Bold = true;
                //Apply fill color for shape.
                rectangle.FillFormat.Fill = true;
                rectangle.FillFormat.Color = Color.LightGray;
                //Set transparency (opacity) to the shape fill color.
                rectangle.FillFormat.Transparency = 75;
                //Apply wrap formats.
                rectangle.WrapFormat.TextWrappingStyle = TextWrappingStyle.Square;
                rectangle.WrapFormat.TextWrappingType = TextWrappingType.Right;
                //Set horizontal and vertical origin.
                rectangle.HorizontalOrigin = HorizontalOrigin.Margin;
                rectangle.VerticalOrigin = VerticalOrigin.Page;
                //Set line format.
                rectangle.LineFormat.DashStyle = LineDashing.Dot;
                rectangle.LineFormat.Color = Color.DarkGray;
                //Set the left internal margin for the shape.
                rectangle.TextFrame.InternalMargin.Left = 30;
                //Set the right internal margin for the shape.
                rectangle.TextFrame.InternalMargin.Right = 24;
                //Set the bottom internal margin for the shape.
                rectangle.TextFrame.InternalMargin.Bottom = 18;
                //Set the top internal margin for the shape.
                rectangle.TextFrame.InternalMargin.Top = 6;
                //Create a file stream.
                using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"../../../Sample.docx"), FileMode.Create, FileAccess.ReadWrite))
                {
                    //Save the Word document to the file stream.
                    document.Save(outputFileStream, FormatType.Docx);
                }
            }
        }
    }
}
