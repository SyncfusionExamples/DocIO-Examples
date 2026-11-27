using Syncfusion.DocIO.DLS;
using Syncfusion.Drawing;
using Syncfusion.OfficeChart;

using (WordDocument document = new WordDocument())
{
    // Add a section
    IWSection section = document.AddSection();

    // Add spacing before inserting chart
    section.AddParagraph();

    // Add chart to the section
    IWParagraph chartParagraph = section.AddParagraph();
    WChart chart = chartParagraph.AppendChart(500, 300);

    // Set chart type to Stacked Bar
    chart.ChartType = OfficeChartType.Bar_Stacked;

    // Chart data
    chart.ChartData.SetValue(1, 1, "Task");
    chart.ChartData.SetValue(1, 2, "Start");
    chart.ChartData.SetValue(1, 3, "Duration");

    chart.ChartData.SetValue(2, 1, "Planning");
    chart.ChartData.SetValue(2, 2, 0);
    chart.ChartData.SetValue(2, 3, 5);

    chart.ChartData.SetValue(3, 1, "Design");
    chart.ChartData.SetValue(3, 2, 5);
    chart.ChartData.SetValue(3, 3, 7);

    chart.ChartData.SetValue(4, 1, "Testing");
    chart.ChartData.SetValue(4, 2, 12);
    chart.ChartData.SetValue(4, 3, 4);

    chart.ChartData.SetValue(5, 1, "Deployment");
    chart.ChartData.SetValue(5, 2, 16);
    chart.ChartData.SetValue(5, 3, 3);

    chart.ChartData.SetValue(6, 1, "Support");
    chart.ChartData.SetValue(6, 2, 19);
    chart.ChartData.SetValue(6, 3, 1);

    // Set data range
    chart.DataRange = chart.ChartData[1, 1, 6, 3];

    // First series = Start (hide transparent)
    IOfficeChartSerie startSeries = chart.Series[0];
    startSeries.SerieFormat.Fill.FillType = OfficeFillType.SolidColor;
    startSeries.SerieFormat.Fill.ForeColor = Color.Transparent;

    // Second series = Duration
    IOfficeChartSerie durationSeries = chart.Series[1];
    durationSeries.SerieFormat.Fill.ForeColor = Color.FromArgb(141, 201, 219);

    // Reverse task order (top to bottom)
    chart.PrimaryCategoryAxis.ReversePlotOrder = true;

    // Show vertical gridlines
    chart.PrimaryValueAxis.HasMajorGridLines = true;

    // Remove legend
    chart.HasLegend = false;

    // Gap width similar to Gantt chart
    durationSeries.SerieFormat.CommonSerieOptions.GapWidth = 30;

    // Chart area background
    chart.ChartArea.Fill.ForeColor = Color.Transparent;

    // Axis min/max
    chart.PrimaryValueAxis.MinimumValue = 0;
    chart.PrimaryValueAxis.MaximumValue = 20;
    chart.PrimaryValueAxis.MajorUnit = 2;

    // Set chart title
    chart.ChartTitle = "Gantt Chart";

    // Hide major gridlines on value axis
    chart.PrimaryValueAxis.HasMajorGridLines = false;

    // Save document
    document.Save(@"../../../Output/GanttChart.docx");
}