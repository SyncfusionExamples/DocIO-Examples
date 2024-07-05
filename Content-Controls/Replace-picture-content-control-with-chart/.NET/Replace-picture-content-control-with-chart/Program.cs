using Syncfusion.DocIO.DLS;
using Syncfusion.DocIO;
using Syncfusion.DocIORenderer;
using Syncfusion.OfficeChart;


// Open the file as stream.
using (FileStream docStream = new FileStream(Path.GetFullPath(@"../../../Data/Input.docx"), FileMode.Open))
{
    // Load file stream into Word document.
    using (WordDocument document = new WordDocument(docStream, Syncfusion.DocIO.FormatType.Automatic))
    {
        string[] propertyNames = { "ContentControlProperties.Title", "ContentControlProperties.Type" };
        string[] propertyValues = { "Chart", "Picture" };
        // Find BlockContentControl by given properties
        BlockContentControl contentControl = document.FindItemByProperties(EntityType.BlockContentControl, propertyNames, propertyValues) as BlockContentControl;

        if (contentControl != null)
        {
            int index = contentControl.OwnerTextBody.ChildEntities.IndexOf(contentControl);

            // Create a new paragraph to hold the chart image.
            WParagraph paragraph = new WParagraph(document);

            // Generate the chart and get it as an image stream.
            Stream chartImage = GenerateChartAndGetChartAsImage();

            // Create a new image instance and load the chart image.
            WPicture picture = (WPicture)paragraph.AppendPicture(chartImage);

            // Set picture dimensions.
            picture.Height = 300;
            picture.Width = 300;

            // To replace Picture content Control insert Picture paragraph and remove the content control.
            contentControl.OwnerTextBody.ChildEntities.Insert(index, paragraph);
            contentControl.OwnerTextBody.ChildEntities.RemoveAt(index + 1);
        }


        // Create file stream for output document.
        using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"../../../Data/Sample.docx"), FileMode.Create, FileAccess.ReadWrite))
        {
            // Save the Word document to the file stream.
            document.Save(outputFileStream, FormatType.Docx);
        }
    }
}


/// <summary>
/// Generates a chart and saves it as an image.
/// </summary>
/// <returns>A stream containing the chart image.</returns>
static Stream GenerateChartAndGetChartAsImage()
{
    // Create a new Word document.
    using (WordDocument document = new WordDocument())
    {
        // Add a section to the document.
        IWSection section = document.AddSection();
        // Add a paragraph to the section.
        IWParagraph paragraph = section.AddParagraph();
        // Create and append the chart to the paragraph.
        WChart chart = paragraph.AppendChart(446, 270);

        // Set chart data.
        chart.ChartData.SetValue(1, 1, "Month");
        chart.ChartData.SetValue(2, 1, "Jan");
        chart.ChartData.SetValue(3, 1, "Feb");
        chart.ChartData.SetValue(4, 1, "Mar");
        chart.ChartData.SetValue(5, 1, "Apr");
        chart.ChartData.SetValue(6, 1, "May");
        chart.ChartData.SetValue(7, 1, "Jun");
        chart.ChartData.SetValue(8, 1, "Jul");
        chart.ChartData.SetValue(9, 1, "Aug");
        chart.ChartData.SetValue(10, 1, "Sep");
        chart.ChartData.SetValue(11, 1, "Oct");
        chart.ChartData.SetValue(12, 1, "Nov");
        chart.ChartData.SetValue(13, 1, "Dec");
        chart.ChartData.SetValue(1, 2, "Rainy Days");
        chart.ChartData.SetValue(2, 2, 12);
        chart.ChartData.SetValue(3, 2, 11);
        chart.ChartData.SetValue(4, 2, 10);
        chart.ChartData.SetValue(5, 2, 9);
        chart.ChartData.SetValue(6, 2, 8);
        chart.ChartData.SetValue(7, 2, 6);
        chart.ChartData.SetValue(8, 2, 4);
        chart.ChartData.SetValue(9, 2, 6);
        chart.ChartData.SetValue(10, 2, 7);
        chart.ChartData.SetValue(11, 2, 8);
        chart.ChartData.SetValue(12, 2, 10);
        chart.ChartData.SetValue(13, 2, 11);
        chart.ChartData.SetValue(1, 3, "Profit");
        chart.ChartData.SetValue(2, 3, 3574);
        chart.ChartData.SetValue(3, 3, 4708);
        chart.ChartData.SetValue(4, 3, 5332);
        chart.ChartData.SetValue(5, 3, 6693);
        chart.ChartData.SetValue(6, 3, 8843);
        chart.ChartData.SetValue(7, 3, 12347);
        chart.ChartData.SetValue(8, 3, 15180);
        chart.ChartData.SetValue(9, 3, 11198);
        chart.ChartData.SetValue(10, 3, 9739);
        chart.ChartData.SetValue(11, 3, 9846);
        chart.ChartData.SetValue(12, 3, 6620);
        chart.ChartData.SetValue(13, 3, 5085);

        // Set region of chart data.
        chart.DataRange = chart.ChartData[1, 1, 13, 3];
        // Set chart series in the column for assigned data region.
        chart.IsSeriesInRows = false;
        // Set chart title.
        chart.ChartTitle = "Combination Chart";

        // Set data labels.
        IOfficeChartSerie series1 = chart.Series[0];
        IOfficeChartSerie series2 = chart.Series[1];
        // Set series type.
        series1.SerieType = OfficeChartType.Column_Clustered;
        series2.SerieType = OfficeChartType.Line;
        series2.UsePrimaryAxis = false;

        // Set data labels.
        series1.DataPoints.DefaultDataPoint.DataLabels.IsValue = true;

        // Set legend.
        chart.HasLegend = true;
        chart.Legend.Position = OfficeLegendPosition.Bottom;

        // Set chart type.
        chart.ChartType = OfficeChartType.Combination_Chart;

        // Set secondary axis on right side.
        chart.SecondaryValueAxis.TickLabelPosition = OfficeTickLabelPosition.TickLabelPosition_High;

        // Create an instance of DocIORenderer.
        using (DocIORenderer renderer = new DocIORenderer())
        {
            // Convert chart to an image.
            Stream stream = chart.SaveAsImage();
            return stream;
        }
    }
}

