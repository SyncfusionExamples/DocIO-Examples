using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.Drawing;
using System.IO;

namespace Add_shapes_in_Word
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
                //Adds new shape to the document.
                Shape rectangle = paragraph.AppendShape(AutoShapeType.RoundedRectangle, 150, 100);
                //Sets position for shape.
                rectangle.VerticalPosition = 72;
                rectangle.HorizontalPosition = 72;
                paragraph = section.AddParagraph() as WParagraph;
                //Adds textbody contents to the shape.
                paragraph = rectangle.TextBody.AddParagraph() as WParagraph;
                IWTextRange text = paragraph.AppendText("This text is in rounded rectangle shape");
                text.CharacterFormat.TextColor = Color.Green;
                text.CharacterFormat.Bold = true;
                //Adds another shape to the document. 
                paragraph = section.AddParagraph() as WParagraph;
                paragraph.AppendBreak(BreakType.LineBreak);
                Shape pentagon = paragraph.AppendShape(AutoShapeType.Pentagon, 100, 100);
                paragraph = pentagon.TextBody.AddParagraph() as WParagraph;
                paragraph.AppendText("This text is in pentagon shape");
                pentagon.HorizontalPosition = 72;
                pentagon.VerticalPosition = 200;
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
