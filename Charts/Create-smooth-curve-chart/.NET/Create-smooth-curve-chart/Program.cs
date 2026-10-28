using Syncfusion.DocIO.DLS;
using Syncfusion.OfficeChart;
using Syncfusion.Drawing;


// Create a new Word document
using (WordDocument document = new WordDocument())
{
    IWSection section = document.AddSection();
    section.PageSetup.PageSize = PageSize.A3;

    // Add title paragraphs
    IWParagraph titlePara1 = section.AddParagraph();
    titlePara1.AppendText("DUKE ENERGY FLORIDA").CharacterFormat.Bold = true;
    titlePara1.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Center;

    IWParagraph titlePara2 = section.AddParagraph();
    titlePara2.AppendText("ACCOUNT 352 STRUCTURES AND IMPROVEMENTS").CharacterFormat.Bold = true;
    titlePara2.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Center;

    IWParagraph titlePara3 = section.AddParagraph();
    titlePara3.AppendText("ORIGINAL AND SMOOTH SURVIVOR CURVES").CharacterFormat.Bold = true;
    titlePara3.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Center;

    // Add paragraph for chart
    IWParagraph chartPara = section.AddParagraph();

    // Create chart
    WChart chart = chartPara.AppendChart(700, 350);
    chart.ChartType = OfficeChartType.Scatter_SmoothedLine;
    chart.ChartTitle = "Smooth Curve chart";

    #region Define Chart Data
    int[] ages = { 0, 10, 20, 30, 40, 50, 60, 70, 80, 90, 100 };
    int[] originalCurve = { 100, 99, 97, 93, 85, 65, 40, 20, 8, 2, 0 };

    // Set up the chart data
    // First row with headers
    chart.ChartData.SetValue(1, 1, "ORIGINAL CURVE");
    chart.ChartData.SetValue(1, 2, "1943-2022 Experience\n1925-2022 Placements");
    chart.ChartData.SetValue(1, 3, "1943-2022 Experience\n1925-2022 Placements");
    chart.ChartData.SetValue(1, 4, "1943-2022 Experience\n1925-2022 Placements");
    chart.ChartData.SetValue(1, 5, "1943-2022 Experience\n1925-2022 Placements");
    chart.ChartData.SetValue(1, 6, "1943-2022 Experience\n1925-2022 Placements");

    // Fill data rows - note we're only filling the original curve column
    for (int i = 0; i < ages.Length; i++)
    {
        chart.ChartData.SetValue(i + 2, 1, ages[i]);
        chart.ChartData.SetValue(i + 2, 3, originalCurve[i]);

        if (i % 2 == 0) // Add some scattered points for other series
        {
            chart.ChartData.SetValue(i + 2, 2, originalCurve[i] - 3);
            chart.ChartData.SetValue(i + 2, 4, originalCurve[i] + 2);
        }
        if (i % 3 == 0)
        {
            chart.ChartData.SetValue(i + 2, 5, originalCurve[i] - 5);
            chart.ChartData.SetValue(i + 2, 6, originalCurve[i] + 4);
        }
    }

    // Set data range
    chart.DataRange = chart.ChartData[1, 1, ages.Length + 1, 6];
    #endregion


    #region Configure Series
    for (int i = 0; i < chart.Series.Count; i++)
    {
        // Set X values to use the ages column
        chart.Series[i].CategoryLabels = chart.ChartData[2, 1, ages.Length + 1, 1];

        // Set Y values to use the respective data column
        chart.Series[i].Values = chart.ChartData[2, i + 2, ages.Length + 1, i + 2];

        // Configure series appearance
        if (i == 1) // Original curve (index 1 corresponds to column 3)
        {
            chart.Series[i].SerieFormat.LineProperties.LinePattern = OfficeChartLinePattern.Solid;
            chart.Series[i].SerieFormat.IsSmoothedLine = true;
            chart.Series[i].SerieFormat.LineProperties.LineWeight = OfficeChartLineWeight.Narrow;
            chart.Series[i].SerieFormat.MarkerStyle = OfficeChartMarkerType.None;
            chart.Series[i].SerieFormat.LineProperties.LineColor = Syncfusion.Drawing.Color.Black;

            chart.Series[i].DataPoints.DefaultDataPoint.DataLabels.IsValue = false;

            IOfficeChartDataPoint point = chart.Series[i].DataPoints[5]; // 0-based index
            point.DataLabels.IsSeriesName = true;
            point.DataLabels.Position = OfficeDataLabelPosition.Above;
            point.DataLabels.Layout.Left = 80;
            point.DataLabels.Layout.Top = 0.5;
            point.DataLabels.ShowLeaderLines = true;

        }
        else
        {
            // For other series, show only markers
            chart.Series[i].SerieFormat.LineProperties.LinePattern = OfficeChartLinePattern.None;

            // Set different marker styles
            switch (i)
            {
                case 0: chart.Series[i].SerieFormat.MarkerStyle = OfficeChartMarkerType.Triangle; break;
                case 2: chart.Series[i].SerieFormat.MarkerStyle = OfficeChartMarkerType.Circle; break;
                case 3: chart.Series[i].SerieFormat.MarkerStyle = OfficeChartMarkerType.Diamond; break;
                case 4: chart.Series[i].SerieFormat.MarkerStyle = OfficeChartMarkerType.Square; break;
                default: chart.Series[i].SerieFormat.MarkerStyle = OfficeChartMarkerType.Star; break;
            }
            chart.Series[i].SerieFormat.MarkerSize = 5;
        }
    }
    #endregion


    #region Configure Chart Axes and Legend
    chart.Legend.Position = OfficeLegendPosition.Top;

    // Add vertical gridlines
    chart.PrimaryCategoryAxis.HasMajorGridLines = true;
    chart.PrimaryCategoryAxis.MajorGridLines.Border.LinePattern = OfficeChartLinePattern.Solid;
    chart.PrimaryCategoryAxis.MajorGridLines.Border.LineWeight = OfficeChartLineWeight.Hairline;
    chart.PrimaryCategoryAxis.MajorGridLines.Border.LineColor = Color.Black;

    chart.PrimaryCategoryAxis.TickLabelSpacing = 10;

    // Configure the value axis (Y-axis)
    chart.PrimaryValueAxis.MinimumValue = 0;
    chart.PrimaryValueAxis.MaximumValue = 100;
    chart.PrimaryValueAxis.MajorUnit = 10;
    chart.PrimaryValueAxis.HasMajorGridLines = true;
    chart.PrimaryValueAxis.MajorGridLines.Border.LinePattern = OfficeChartLinePattern.Solid;
    chart.PrimaryValueAxis.MajorGridLines.Border.LineWeight = OfficeChartLineWeight.Hairline;
    chart.PrimaryValueAxis.MajorGridLines.Border.LineColor = Color.Black;
    chart.PrimaryValueAxis.Title = "PERCENT SURVIVING";

    // Configure the category axis (X-axis)
    chart.PrimaryCategoryAxis.MinimumValue = 0;
    chart.PrimaryCategoryAxis.MaximumValue = 120;
    chart.PrimaryCategoryAxis.MajorUnit = 10;
    chart.PrimaryCategoryAxis.Title = "AGE IN YEARS";
    #endregion

    document.Save(@"../../../Output/SurvivorCurves.docx");
}
