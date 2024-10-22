using Syncfusion.DocIO.DLS;

// Creates a new instance of WordDocument (Empty Word Document).
using (WordDocument document = new WordDocument())
{
    // Adds a section to the document.
    IWSection sec = document.AddSection();
    // Adds a paragraph to the section.
    IWParagraph paragraph = sec.AddParagraph();
    // Loads the Excel file as a stream.
    Stream excelStream = File.OpenRead(Path.GetFullPath("Data/InputTemplate.xlsx"));
    // Creates and appends a chart to the paragraph with the Excel stream as a parameter.
    // The chart is created based on the data from the Excel file (range A1:D6), with specified width and height.
    WChart chart = paragraph.AppendChart(excelStream, 1, "A1:D6", 470, 300);
    // Sets the chart type to Stacked Bar Chart.
    chart.ChartType = Syncfusion.OfficeChart.OfficeChartType.Bar_Stacked;
    // Apply chart elements.
    // Sets the chart title.
    chart.ChartTitle = "Stacked Bar Chart";
    // Displays data labels for the third series (Series 2).
    chart.Series[2].DataPoints.DefaultDataPoint.DataLabels.IsValue = true;

    // Manually positions the data labels for each data point in the third series.
    chart.Series[2].DataPoints[0].DataLabels.Text = "Label 1";
    chart.Series[2].DataPoints[0].DataLabels.Layout.ManualLayout.Left = 0.1;
    chart.Series[2].DataPoints[1].DataLabels.Text = "Label 2";
    chart.Series[2].DataPoints[1].DataLabels.Layout.ManualLayout.Left = 0.1;
    chart.Series[2].DataPoints[2].DataLabels.Text = "Label 3";
    chart.Series[2].DataPoints[2].DataLabels.Layout.ManualLayout.Left = 0.13;
    chart.Series[2].DataPoints[3].DataLabels.Text = "Label 4";
    chart.Series[2].DataPoints[3].DataLabels.Layout.ManualLayout.Left = 0.18;
    chart.Series[2].DataPoints[4].DataLabels.Text = "Label 5";
    chart.Series[2].DataPoints[4].DataLabels.Layout.ManualLayout.Left = 0.20;

    // Sets the chart legend and positions it at the bottom of the chart.
    chart.HasLegend = true;
    chart.Legend.Position = Syncfusion.OfficeChart.OfficeLegendPosition.Bottom;
    using (FileStream outputStream = new FileStream(Path.GetFullPath("Output/Result.docx"), FileMode.Create, FileAccess.Write))
    {
        // Saves the generated Word document to the specified file stream in DOCX format.
        document.Save(outputStream, Syncfusion.DocIO.FormatType.Docx);
    }
}

