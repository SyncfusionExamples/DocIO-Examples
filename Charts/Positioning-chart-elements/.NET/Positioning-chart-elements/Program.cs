using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.OfficeChart;
using System.Collections.Generic;
using System.IO;

namespace Positioning_chart_elements
{
    class Program
    {
        static void Main(string[] args)
        {
            //Creates a new word document.
            using (WordDocument document = new WordDocument())
            {
                //Adds section to the document.
                IWSection sec = document.AddSection();
                //Adds paragraph to the section.
                IWParagraph paragraph = sec.AddParagraph();
                //Creates and Appends chart to the paragraph.
                WChart chart = paragraph.AppendChart(470, 300);
                //Inputs data for chart.
                List<BarChartData> dataList = new List<BarChartData>();
                BarChartData column1 = new BarChartData("P1", 286, 1300);
                BarChartData column2 = new BarChartData("P2", 680, 700);
                BarChartData column3 = new BarChartData("P3", 288, 1280);
                BarChartData column4 = new BarChartData("P4", 200, 1200);
                BarChartData column5 = new BarChartData("P5", 731, 2660);
                dataList.Add(column1);
                dataList.Add(column2);
                dataList.Add(column3);
                dataList.Add(column4);
                dataList.Add(column5);
                //Sets chart data by using IEnumerable overload.
                chart.SetDataRange(dataList, 1, 1);
                //Sets chart type and title.
                chart.ChartTitle = "Purchase Details";
                chart.ChartArea.Border.LinePattern = OfficeChartLinePattern.None;
                //Axis titles.
                chart.PrimaryCategoryAxis.Title = "Products";
                chart.PrimaryValueAxis.Title = "In Dollars";
                //Sets position for plot area.
                chart.PlotArea.Layout.LeftMode = LayoutModes.auto;
                chart.PlotArea.Layout.TopMode = LayoutModes.factor;
                chart.PlotArea.Layout.LayoutTarget = LayoutTargets.outer;
                //Sets position for title area.
                chart.ChartTitleArea.Layout.Left = 10;
                chart.ChartTitleArea.Layout.Top = 8;
                //Sets position for chart legend.
                chart.Legend.Layout.LeftMode = LayoutModes.factor;
                chart.Legend.Layout.TopMode = LayoutModes.edge;
                //Creates file stream.
                using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"Output/Output.docx"), FileMode.Create, FileAccess.ReadWrite))
                {
                    //Saves the Word document to file stream.
                    document.Save(outputFileStream, FormatType.Docx);
                }
            }
        }
    }

    public class BarChartData
    {
        string name;
        int purchase;
        int expense;
        public string Name
        {
            get
            {
                return name;
            }
            set
            {
                name = value;
            }
        }
        public int Purchase
        {
            get
            {
                return purchase;
            }
            set
            {
                purchase = value;
            }
        }
        public int Expense
        {
            get
            {
                return expense;
            }
            set
            {
                expense = value;
            }
        }
        public BarChartData(string name, int purchase, int expense)
        {
            Name = name;
            Purchase = purchase;
            Expense = expense;
        }
    }
}
