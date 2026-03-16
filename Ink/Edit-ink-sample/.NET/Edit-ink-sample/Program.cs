using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.Drawing;
using Syncfusion.Office;

namespace Edit_ink
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
                    // Access the first section of the document.
                    WSection section = document.Sections[0];
                    // Access the first ink and customize its trace points.
                    WInk firstInk = section.Paragraphs[0].ChildEntities[0] as WInk;
                    // Move the ink vertically.
                    firstInk.VerticalPosition = 25f;
                    // Copy existing points into the new array.
                    int oldTracePointsLength = firstInk.Traces[0].Points.Length;
                    int newTracePointsLength = oldTracePointsLength + 3;
                    PointF[] newTracePoints = new PointF[newTracePointsLength];
                    PointF[] oldTracePoints = firstInk.Traces[0].Points;
                    Array.Copy(oldTracePoints, newTracePoints, oldTracePointsLength);
                    newTracePoints[newTracePoints.Length - 3] = new PointF(oldTracePoints[3].X, 0);
                    newTracePoints[newTracePoints.Length - 2] = new PointF(oldTracePoints[0].X, 0);
                    newTracePoints[newTracePoints.Length - 1] = new PointF(oldTracePoints[0].X, oldTracePoints[0].Y);
                    // Update the trace points of the first ink with the new array.
                    firstInk.Traces[0].Points = newTracePoints;

                    // Access the second ink and customize its brush effect.
                    WInk secondInk = section.Paragraphs[1].ChildEntities[0] as WInk;
                    IOfficeInkTrace secondInkTrace = secondInk.Traces[0];
                    // Set the ink size (thickness) to 1 point.
                    secondInkTrace.Brush.Size = new SizeF(1f, 1f);

                    // Access the third ink and customize its container width.
                    WInk thirdInk = section.Paragraphs[2].ChildEntities[0] as WInk;
                    // Set the width of the ink container to 130 points.
                    thirdInk.Width = 130f;

                    // Access the fourth ink and customize its brush color.
                    WParagraph paragraph = section.Tables[0].Rows[0].Cells[0].ChildEntities[0] as WParagraph;
                    WInk fourthInk = paragraph.ChildEntities[0] as WInk;
                    IOfficeInkTrace fourthInkTrace = fourthInk.Traces[0];
                    // Set the color of the ink stroke to Yellow.
                    fourthInkTrace.Brush.Color = Color.Yellow;
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

