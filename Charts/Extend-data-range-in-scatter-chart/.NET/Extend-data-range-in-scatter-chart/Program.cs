using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.OfficeChart;


//Open an existing document
using (WordDocument document = new WordDocument(@"../../../Data/Template.docx", FormatType.Docx))
{
    // Access the chart from the document (e.g., first child entity of the last paragraph)
    WChart chart = document.LastParagraph.ChildEntities[0] as WChart;
    chart.PrimaryCategoryAxis.Title = "Experience (Years)";
    chart.PrimaryValueAxis.Title = "Salary (₹ Lakhs)";
    if (chart != null)
    {
        // Add the two new data rows to ChartData
        chart.ChartData.SetValue(5, 1, 7);
        chart.ChartData.SetValue(5, 2, 10.0);

        // Row 6: 50, 500
        chart.ChartData.SetValue(6, 1, 8);
        chart.ChartData.SetValue(6, 2, 12.5);

        // Update the series data ranges to reflect the extended range
        if (chart.Series.Count > 0)
        {
            IOfficeChartSerie series = chart.Series[0];
            // X values (Category labels): Column 1, Rows 2 to 6
            chart.PrimaryCategoryAxis.CategoryLabels = chart.ChartData[2, 1, 6, 1];
            // Y values (Values): Column 2, Rows 2 to 6
            series.Values = chart.ChartData[2, 2, 6, 2];
        }
        // Refresh the chart to re-evaluate and apply the extended data
        chart.Refresh();
    }
    //Save the Word file.
    document.Save(@"../../../Output/Output.docx", FormatType.Docx);
}