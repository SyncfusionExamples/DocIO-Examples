using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.OfficeChart;
using System.IO;

//Create a new Word document.
using (WordDocument document = new WordDocument())
{
    //Add a section to the document.
    IWSection section = document.AddSection();
    //Add a paragraph to the section.
    IWParagraph paragraph = section.AddParagraph();
    //Create and append the chart to the paragraph.
    WChart chart = paragraph.AppendChart(446, 270);
    //Set chart type.
    chart.ChartType = OfficeChartType.Line;

    //Add chart data (Example)
    chart.ChartData.SetValue(1, 1, "Month");
    chart.ChartData.SetValue(1, 2, "Sales");
    chart.ChartData.SetValue(2, 1, "January");
    chart.ChartData.SetValue(2, 2, 1000);
    chart.ChartData.SetValue(3, 1, "February");
    chart.ChartData.SetValue(3, 2, 2000);
    chart.ChartData.SetValue(4, 1, "March");
    chart.ChartData.SetValue(4, 2, 1500);
    chart.ChartData.SetValue(5, 1, "April");
    chart.ChartData.SetValue(5, 2, 500);

    //Create chart series
    IOfficeChartSerie series = chart.Series.Add("Sales");

    //Enable smooth line for the series
    series.SerieFormat.IsSmoothedLine = true;
    //Set region of Chart data.
    chart.DataRange = chart.ChartData[1, 1, 5, 2];

    //Set legend.
    chart.HasLegend = true;
    chart.Legend.Position = OfficeLegendPosition.Bottom;
    //Save the Word document to the file stream.
    document.Save(@"../../../Output/Result.docx", FormatType.Docx);
}