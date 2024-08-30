using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.OfficeChart;
using System;
using System.ComponentModel;
using System.IO;

namespace Find_and_replace_text_with_chart
{
    internal class Program
    {
        static void Main(string[] args)
        {
            using (FileStream fileStreamPath = new FileStream(Path.GetFullPath(@"Data/Input.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                //Opens an existing Word document.
                using (WordDocument document = new WordDocument(fileStreamPath, FormatType.Docx))
                {
                    TextBodyPart bodyPart = new TextBodyPart(document);
                    //Create new paragraph
                    WParagraph paragraph= CreateBarChart(document);    
                    bodyPart.BodyItems.Add(paragraph);
                    //Replaces the placeholder text with a new chart.
                    document.Replace("[Purchase details]", bodyPart, true, true, true);
                    //Creates file stream.
                    using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"Output/Result.docx"), FileMode.Create, FileAccess.ReadWrite))
                    {
                        //Saves the Word document to file stream.
                        document.Save(outputFileStream, FormatType.Docx);
                    }
                }
            }
        } 
        /// <summary>
        /// Create a bar chart inside a new paragraph.
        /// </summary>
        /// <param name="document"></param>
        /// <returns></returns>
        private static WParagraph CreateBarChart(WordDocument document)
        {
            //Create new paragraph
            WParagraph paragraph = new WParagraph(document);
            //Append new chart
            WChart barchart = paragraph.AppendChart(400, 300);
            //Set type as bar chart
            barchart.ChartType = OfficeChartType.Bar_Clustered;
            //Add data for the bar chart
            AddBarChartData(barchart);

            //Set the other parts of the chart.
            barchart.ChartTitle = "Purchase Details";
            barchart.ChartTitleArea.FontName = "Calibri";
            barchart.ChartTitleArea.Size = 14;
            IOfficeChartSerie serie1 = barchart.Series.Add("Sum of Future Expenses");
            serie1.Values = barchart.ChartData[2, 2, 6, 2];
            IOfficeChartSerie serie2 = barchart.Series.Add("Sum of Purchases");
            serie2.Values = barchart.ChartData[2, 3, 6, 3];

            barchart.HasDataTable = true;
            barchart.DataTable.HasBorders = true;
            barchart.DataTable.HasHorzBorder = true;
            barchart.DataTable.HasVertBorder = true;
            barchart.DataTable.ShowSeriesKeys = true;
            barchart.HasLegend = false;

            barchart.PlotArea.Border.LinePattern = OfficeChartLinePattern.Solid;
            barchart.ChartArea.Border.LinePattern = OfficeChartLinePattern.Solid;
            barchart.PrimaryCategoryAxis.CategoryLabels = barchart.ChartData[2, 1, 6, 1];
            barchart.PrimaryCategoryAxis.Font.Size = 12;
            barchart.PrimaryCategoryAxis.MajorTickMark = OfficeTickMark.TickMark_None;
            barchart.PrimaryValueAxis.MajorTickMark = OfficeTickMark.TickMark_None;

            //Return the new paragraph that has bar chart in it.
            return paragraph;
        }
        /// <summary>
        /// Add the data for bar chart
        /// </summary>
        /// <param name="chart"></param>
        private static void AddBarChartData(WChart chart)
        {
            chart.ChartData.SetValue(1, 2, "Sum of Future Expenses");
            chart.ChartData.SetValue(1, 3, "Sum of Purchases");
            chart.ChartData.SetValue(2, 1, "Nancy Davalio");
            chart.ChartData.SetValue(2, 2, 1300);
            chart.ChartData.SetValue(2, 3, 600);
            chart.ChartData.SetValue(3, 1, "Andrew Fuller");
            chart.ChartData.SetValue(3, 2, 680);
            chart.ChartData.SetValue(3, 3, 1000);
            chart.ChartData.SetValue(4, 1, "Janet Leverling");
            chart.ChartData.SetValue(4, 2, 1280);
            chart.ChartData.SetValue(4, 3, 800);
            chart.ChartData.SetValue(5, 1, "Margaret Peacock");
            chart.ChartData.SetValue(5, 2, 2000);
            chart.ChartData.SetValue(5, 3, 400);
            chart.ChartData.SetValue(6, 1, "Steven Buchanan");
            chart.ChartData.SetValue(6, 2, 2660);
            chart.ChartData.SetValue(6, 3, 731);
        }

    }
}
