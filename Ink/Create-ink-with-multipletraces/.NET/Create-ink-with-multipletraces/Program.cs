using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.Drawing;
using Syncfusion.Office;

namespace Create_ink_with_multipletraces
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
                IWTextRange firstText = paragraph.AppendText("Created Ink with multiple traces");
                //Apply formatting for first text range
                firstText.CharacterFormat.FontSize = 14;
                firstText.CharacterFormat.Bold = true;
                //Adds new ink to the document.
                WInk ink = paragraph.AppendInk(450, 350);
                // Sets the horizontal position of the ink object.
                ink.HorizontalPosition = 30;
                // Sets the Vertical position of the ink object.
                ink.VerticalPosition = 50;
                // Sets the text wrapping style for the ink object to be in front of text.
                ink.WrapFormat.TextWrappingStyle = TextWrappingStyle.InFrontOfText;
                // Gets the ink traces collection from the ink object.
                IOfficeInkTraces traces = ink.Traces;
                // Gets all trace point arrays from the helper method.
                List<PointF[]> pointsCollection = GetPoints();
                // Adds each trace to the ink object.
                foreach (var points in pointsCollection)
                {
                    // Adds the trace to the ink object.
                    IOfficeInkTrace trace = traces.Add(points);
                    // Sets the brush color for the trace to red.
                    trace.Brush.Color = Color.Red;
                    // Sets the brush size for the ink stroke.
                    trace.Brush.Size = new SizeF(5f, 5f);
                }
                using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"Output/Result.docx"), FileMode.Create, FileAccess.ReadWrite))
                {
                    //Saves the Word document to file stream.
                    document.Save(outputFileStream, FormatType.Docx);
                }
            }
        }
        /// <summary>
        /// A collection where each <see cref="PointF"/> array represents a single stroke.
        /// </summary>
        static List<PointF[]> GetPoints()
        {
            return new List<PointF[]>
            {
                //Trace_i
                new PointF[] {
                    new PointF(20f, 10f),
                    new PointF(20f, 140f),
                },
                //Trace_n
                new PointF[]
                {
                    new PointF(60f, 80f),
                    new PointF(60f, 100f),
                    new PointF(60f, 140f),
                    new PointF(60f, 92f),
                    new PointF(70f, 86f),
                    new PointF(88f, 84f),
                    new PointF(100f, 92f),
                    new PointF(106f, 108f),
                    new PointF(110f, 140f)
                },
                //Trace_k
                new PointF[] {
                    new PointF(140f, 10f),
                    new PointF(140f, 140f),
                    new PointF(140f, 80f),
                    new PointF(180f, 20f),
                    new PointF(140f, 80f),
                    new PointF(180f, 140f)
                }
            };
        }
    }
}

