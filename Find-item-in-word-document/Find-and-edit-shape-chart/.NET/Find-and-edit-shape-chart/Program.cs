using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;

namespace Find_and_edit_shape_chart
{
    class Program
    {
        static void Main(string[] args)
        {
            using (FileStream fileStreamPath = new FileStream(Path.GetFullPath(@"Data/Template.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                //Opens an existing Word document.
                using (WordDocument document = new WordDocument(fileStreamPath, FormatType.Automatic))
                {
                    //Find shape by name.
                    Shape shape = document.FindItemByProperty(EntityType.AutoShape, "Name", "Adventure Shape") as Shape;
                    //Resize the shape.  
                    if (shape != null)
                    {
                        shape.Height = 75;
                        shape.Width = 200;
                    }
                    //Find chart by name.
                    WChart chart = document.FindItemByProperty(EntityType.Chart, "Name", "Adventure Chart") as WChart;
                    //Set the chart name.  
                    if (chart != null)
                        chart.ChartTitle = "Column chart";

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
}

