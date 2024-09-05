using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.Drawing;
using Syncfusion.OfficeChart;
using System.IO;

namespace Collection_of_shapes_into_group_shape
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
                //Creates paragraph item collections to add child shapes.
                ParagraphItem[] paragraphItems = new ParagraphItem[3];
                //Creates new shape.
                Shape shape = new Shape(document, AutoShapeType.RoundedRectangle);
                //Sets height and width for shape.
                shape.Height = 100;
                shape.Width = 150;
                //Sets Wrapping style for shape.
                shape.WrapFormat.TextWrappingStyle = TextWrappingStyle.InFrontOfText;
                //Sets horizontal and vertical position for shape.
                shape.HorizontalPosition = 7;
                shape.VerticalPosition = 72;
                //Sets horizontal and vertical origin for shape.
                shape.HorizontalOrigin = HorizontalOrigin.Page;
                shape.VerticalOrigin = VerticalOrigin.Page;
                //Sets the shape as paragraph item.
                paragraphItems[0] = shape;
                //Appends new textbox to the document.
                WTextBox textbox = new WTextBox(document);
                //Sets height and width for textbox.
                textbox.TextBoxFormat.Width = 150;
                textbox.TextBoxFormat.Height = 75;
                //Adds new text to the textbox body.
                IWParagraph textboxParagraph = textbox.TextBoxBody.AddParagraph();
                //Adds new text to the textbox paragraph.
                textboxParagraph.AppendText("Text inside text box");
                //Sets wrapping style for textbox.
                textbox.TextBoxFormat.TextWrappingStyle = TextWrappingStyle.Behind;
                //Sets horizontal and vertical position for textbox.
                textbox.TextBoxFormat.HorizontalPosition = 200;
                textbox.TextBoxFormat.VerticalPosition = 200;
                //Sets horizontal and vertical origin for textbox.
                textbox.TextBoxFormat.VerticalOrigin = VerticalOrigin.Page;
                textbox.TextBoxFormat.HorizontalOrigin = HorizontalOrigin.Page;
                //Sets the textbox as paragraph item.
                paragraphItems[1] = textbox;
                //Appends new chart to the document.
                WChart chart = new WChart(document);
                //Sets height and width for chart.
                chart.Height = 270;
                chart.Width = 446;
                //Sets wrapping style for chart.
                chart.WrapFormat.TextWrappingStyle = TextWrappingStyle.InFrontOfText;
                //Sets chart type.
                chart.ChartType = OfficeChartType.Pie;
                chart.VerticalPosition = 350;
                //Sets chart title.
                chart.ChartTitle = "Best Selling Products";
                //Sets font and size for chart title.
                chart.ChartTitleArea.FontName = "Calibri";
                chart.ChartTitleArea.Size = 14;
                //Sets data for chart.
                chart.ChartData.SetValue(1, 1, "");
                chart.ChartData.SetValue(1, 2, "Sales");
                chart.ChartData.SetValue(2, 1, "Phyllis Lapin");
                chart.ChartData.SetValue(2, 2, 141.396);
                chart.ChartData.SetValue(3, 1, "Stanley Hudson");
                chart.ChartData.SetValue(3, 2, 80.368);
                chart.ChartData.SetValue(4, 1, "Bernard Shah");
                chart.ChartData.SetValue(4, 2, 71.155);
                chart.ChartData.SetValue(5, 1, "Patricia Lincoln");
                chart.ChartData.SetValue(5, 2, 47.234);
                chart.ChartData.SetValue(6, 1, "Camembert Pierrot");
                chart.ChartData.SetValue(6, 2, 46.825);
                chart.ChartData.SetValue(7, 1, "Thomas Hardy");
                chart.ChartData.SetValue(7, 2, 42.593);
                chart.ChartData.SetValue(8, 1, "Hanna Moos");
                chart.ChartData.SetValue(8, 2, 41.819);
                chart.ChartData.SetValue(9, 1, "Alice Mutton");
                chart.ChartData.SetValue(9, 2, 32.698);
                chart.ChartData.SetValue(10, 1, "Christina Berglund");
                chart.ChartData.SetValue(10, 2, 29.171);
                chart.ChartData.SetValue(11, 1, "Elizabeth Lincoln");
                chart.ChartData.SetValue(11, 2, 25.696);
                //Creates a new chart series with the name “Sales”.
                IOfficeChartSerie pieSeries = chart.Series.Add("Sales");
                //Sets value for the chart series.
                pieSeries.Values = chart.ChartData[2, 2, 11, 2];
                //Sets data label.
                pieSeries.DataPoints.DefaultDataPoint.DataLabels.IsValue = true;
                pieSeries.DataPoints.DefaultDataPoint.DataLabels.Position = OfficeDataLabelPosition.Outside;
                //Sets background color.
                chart.ChartArea.Fill.ForeColor = Color.FromArgb(242, 242, 242);
                chart.PlotArea.Fill.ForeColor = Color.FromArgb(242, 242, 242);
                chart.ChartArea.Border.LinePattern = OfficeChartLinePattern.None;
                //Sets category labels.
                chart.PrimaryCategoryAxis.CategoryLabels = chart.ChartData[2, 1, 11, 1];
                //Sets the chart as paragraph item.
                paragraphItems[2] = chart;
                //Creates new group shape.
                GroupShape groupShape = new GroupShape(document, paragraphItems);
                groupShape.HorizontalPosition = 72;
                //Adds the group shape to the paragraph.
                paragraph.ChildEntities.Add(groupShape);
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
