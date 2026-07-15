using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.OfficeChart;
using System.IO;


//Loads the template document.
WordDocument document = new WordDocument();
// Adds section to the document.
IWSection sec = document.AddSection();
//Adds paragraph to the section.
IWParagraph paragraph = sec.AddParagraph();
//Creates and Appends chart to the paragraph.
WChart chart = paragraph.AppendChart(446, 270);
chart.ChartType = OfficeChartType.Column_Clustered;
//Assign data
AddChartData(chart);
chart.IsSeriesInRows = false;
//Apply chart elements
//Set chart title
chart.ChartTitle = "Sales Report in Clustered Column Chart";
//Set Datalabels
IOfficeChartSerie serie1 = chart.Series.Add("Amount(in $)");
//Sets the data range of chart series – start row, start column, end row, end column
serie1.Values = chart.ChartData[2, 2, 6, 2];
IOfficeChartSerie serie2 = chart.Series.Add("Count");
//Sets the data range of chart series start row, start column, end row, end column
serie2.Values = chart.ChartData[2, 3, 6, 3];
//Sets the data range of the category axis
chart.PrimaryCategoryAxis.CategoryLabels = chart.ChartData[2, 1, 6, 1];
serie1.DataPoints.DefaultDataPoint.DataLabels.IsValue = true;
serie2.DataPoints.DefaultDataPoint.DataLabels.IsValue = true;
serie1.DataPoints.DefaultDataPoint.DataLabels.Position = OfficeDataLabelPosition.Outside;
serie2.DataPoints.DefaultDataPoint.DataLabels.Position = OfficeDataLabelPosition.Outside;
//Set legend
chart.HasLegend = true;
chart.Legend.Position = OfficeLegendPosition.Bottom;
//Create a file stream.
//Create a file stream.
using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"Output/Output.docx"), FileMode.Create, FileAccess.ReadWrite))
{
    //Save the Word document to the file stream.
    document.Save(outputFileStream, FormatType.Docx);
}


/// <summary>
/// Set the values for the chart
/// </summary>
/// <param name="chart">Represent the instance of the chart</param>

static void AddChartData(WChart chart)
{
    //Set the value for chart data
    chart.ChartData.SetValue(1, 1, "Items");
    chart.ChartData.SetValue(1, 2, "Amount(in $)");
    chart.ChartData.SetValue(1, 3, "Count");

    chart.ChartData.SetValue(2, 1, "Beverages");
    chart.ChartData.SetValue(2, 2, 2776);
    chart.ChartData.SetValue(2, 3, 925);

    chart.ChartData.SetValue(3, 1, "Condiments");
    chart.ChartData.SetValue(3, 2, 1077);
    chart.ChartData.SetValue(3, 3, 378);

    chart.ChartData.SetValue(4, 1, "Confections");
    chart.ChartData.SetValue(4, 2, 2287);
    chart.ChartData.SetValue(4, 3, 880);

    chart.ChartData.SetValue(5, 1, "Dairy Products");
    chart.ChartData.SetValue(5, 2, 1368);
    chart.ChartData.SetValue(5, 3, 581);

    chart.ChartData.SetValue(6, 1, "Grains/Cereals");
    chart.ChartData.SetValue(6, 2, 3325);
    chart.ChartData.SetValue(6, 3, 189);
}