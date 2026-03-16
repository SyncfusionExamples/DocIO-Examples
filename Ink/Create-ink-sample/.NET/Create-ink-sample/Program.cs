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
            using (FileStream fileStreamPath = new FileStream(Path.GetFullPath(@"Data/Template.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                //Open a existing Word document.
                using (WordDocument document = new WordDocument(fileStreamPath, FormatType.Docx))
                {
                    // Get the last paragraph of document.
                    WParagraph paragraph = document.LastParagraph;
                    // Append an ink object to the paragraph.
                    WInk inkObj = paragraph.AppendInk(80, 20);
                    // Get the traces collection from the ink object (traces represent the drawing strokes).
                    IOfficeInkTraces traces = inkObj.Traces;
                    // Retrieve an array of points that define the path of the ink stroke.
                    PointF[] tracePoints = new PointF[] 
                    {
                      new PointF(15f,35f), new PointF(32f,14f), new PointF(42f,12f), new PointF(52f,28f), new PointF(46f,45f),
                      new PointF(52f,36f), new PointF(67f,40f), new PointF(69f,48f), new PointF(61f,42f), new PointF(81f,40f),
                      new PointF(88f,52f), new PointF(107f,38f), new PointF(125f,45f), new PointF(138f,54f), new PointF(123f,49f),
                      new PointF(133f,25f), new PointF(170f,43f), new PointF(190f,47f), new PointF(85f,56f), new PointF(8f,44f)
                    };
                    // Add a new trace (stroke) to the traces collection using the retrieved points.
                    IOfficeInkTrace trace = traces.Add(tracePoints);
                    // Configure the appearance of the ink.
                    // Get the brush object associated with the trace to configure its appearance.
                    IOfficeInkBrush brush = trace.Brush;
                    // Set the ink effect type to None (Pen effect applied).
                    brush.InkEffect = OfficeInkEffectType.None;
                    // Set the color of the ink stroke.
                    brush.Color = Color.Black;
                    // Set the size (thickness) of the ink stroke to 1.5 units.
                    brush.Size = new SizeF(1.5f, 1.5f);
                    using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"Output/Result.docx"), FileMode.Create, FileAccess.ReadWrite))
                    {
                        //Saves the Word document to file stream.
                        document.Save(outputFileStream, FormatType.Docx);
                    }
                }
            }
        }      
    }
}
