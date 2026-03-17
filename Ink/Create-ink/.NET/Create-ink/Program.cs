using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.Drawing;
using Syncfusion.Office;

namespace Create_ink
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
                //Adds new text to the paragraph
                IWTextRange firstText = paragraph.AppendText("Created a triangle using Ink");
                //Apply formatting for first text range
                firstText.CharacterFormat.FontSize = 14;
                firstText.CharacterFormat.Bold = true;
                //Adds new ink to the document.
                WInk ink = paragraph.AppendInk(400, 300);
                // Gets the ink traces collection from the ink object.
                IOfficeInkTraces traces = ink.Traces;
                // Adds new ink stroke with required trace points
                PointF[] triangle = new PointF[] { new PointF(0f, 300f), new PointF(200f, 0f), new PointF(400f, 300f), new PointF(0f, 300f) };
                // Adds a new ink trace to the ink object using the triangle points.
                IOfficeInkTrace trace = traces.Add(triangle);
                // Modify the brush effects and size
                IOfficeInkBrush brush = trace.Brush;
                // Sets the brush size for the ink stroke.
                brush.Size = new SizeF(5f, 5f);
                // Sets the ink effect to 'Galaxy'.
                brush.InkEffect = OfficeInkEffectType.Galaxy;
                using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"Output/Result.docx"), FileMode.Create, FileAccess.ReadWrite))
                {
                    //Saves the Word document to file stream.
                    document.Save(outputFileStream, FormatType.Docx);
                }
            }
        }
    }
}
