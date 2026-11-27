using Syncfusion.DocIO.DLS;
using Syncfusion.OfficeChart;
using Syncfusion.XlsIO;

class Program
{
    static void Main(string[] args)
    {
        ConvertExcelChartToWordChart(@"../../../Data/Input.xlsx", @"../../../Output/Output.docx");
    }
    public static void ConvertExcelChartToWordChart(string excelFilePath, string outputPath)
    {
        // Load the Excel document
        ExcelEngine excelEngine = new ExcelEngine();
        IApplication application = excelEngine.Excel;
        IWorkbook workbook = application.Workbooks.Open(excelFilePath);
        IWorksheet worksheet = workbook.Worksheets[0];

        // Get the first chart from Excel
        IChart excelChart = worksheet.Charts[0];

        // Create Word document
        WordDocument wordDocument = new WordDocument();

        // Create a new section or use existing
        WSection section = wordDocument.LastSection as WSection;
        if (section == null)
        {
            section = wordDocument.AddSection() as WSection;
        }

        // Create a paragraph to add the chart
        IWParagraph paragraph = section.AddParagraph();

        // Create a new chart in Word with the same chart type
        WChart wordChart = paragraph.AppendChart(500, 400);
        wordChart.ChartType = ConvertExcelChartTypeToWord(excelChart.ChartType);

        // Set chart size
        wordChart.Width = (float)excelChart.Width;
        wordChart.Height = (float)excelChart.Height;

        // Copy chart data
        CopyChartData(excelChart, wordChart);

        // Apply formatting
        ApplyChartFormatting(excelChart, wordChart);

        // Save the Word document
        wordDocument.Save(outputPath);
        wordDocument.Close();

        // Close Excel
        workbook.Close();
        excelEngine.Dispose();
    }

    private static OfficeChartType ConvertExcelChartTypeToWord(ExcelChartType excelChartType)
    {
        switch (excelChartType)
        {
            case ExcelChartType.Column_Clustered:
                return OfficeChartType.Column_Clustered;

            case ExcelChartType.Column_Stacked:
                return OfficeChartType.Column_Stacked;

            case ExcelChartType.Line:
                return OfficeChartType.Line;

            case ExcelChartType.Pie:
                return OfficeChartType.Pie;

            case ExcelChartType.Bar_Clustered:
                return OfficeChartType.Bar_Clustered;

            case ExcelChartType.Area:
                return OfficeChartType.Area;

            default:
                return OfficeChartType.Column_Clustered; // Default fallback
        }
    }
    private static void CopyChartData(IChart excelChart, WChart wordChart)
    {
        IRange dataRange = excelChart.DataRange;

        int rowCount = dataRange.Rows.Length;
        int colCount = dataRange.Columns.Length;

        for (int i = 0; i < rowCount; i++)
        {
            for (int j = 0; j < colCount; j++)
            {
                object cellValue = dataRange[i + 1, j + 1].Value;

                if (cellValue != null)
                {
                    wordChart.ChartData.SetValue(i + 1, j + 1, cellValue);
                }
                else
                {
                    wordChart.ChartData.SetValue(i + 1, j + 1, string.Empty);
                }
            }
        }

        wordChart.DataRange = wordChart.ChartData[
            dataRange.Row,
            dataRange.Column,
            dataRange.LastRow,
            dataRange.LastColumn];
    }

    private static void ApplyChartFormatting(IChart excelChart, WChart wordChart)
    {
        // Set chart title
        if (excelChart.HasTitle && !string.IsNullOrWhiteSpace(excelChart.ChartTitle))
        {
            wordChart.ChartTitle = excelChart.ChartTitle;
            wordChart.ChartTitleArea.FontName = excelChart.ChartTitleArea.FontName;
        }

        // Set legend properties
        wordChart.HasLegend = excelChart.HasLegend;

        if (excelChart.HasLegend &&
            excelChart.Legend != null &&
            wordChart.Legend != null)
        {
            switch (excelChart.Legend.Position)
            {
                case ExcelLegendPosition.Bottom:
                    wordChart.Legend.Position = OfficeLegendPosition.Bottom;
                    break;

                case ExcelLegendPosition.Left:
                    wordChart.Legend.Position = OfficeLegendPosition.Left;
                    break;

                case ExcelLegendPosition.Right:
                    wordChart.Legend.Position = OfficeLegendPosition.Right;
                    break;

                case ExcelLegendPosition.Top:
                    wordChart.Legend.Position = OfficeLegendPosition.Top;
                    break;
            }
        }

        // Set value axis title
        if (excelChart.PrimaryValueAxis != null &&
            !string.IsNullOrEmpty(excelChart.PrimaryValueAxis.Title))
        {
            wordChart.PrimaryValueAxis.Title =
                excelChart.PrimaryValueAxis.Title;
        }

        // Set category axis title
        if (excelChart.PrimaryCategoryAxis != null &&
            !string.IsNullOrEmpty(excelChart.PrimaryCategoryAxis.Title))
        {
            wordChart.PrimaryCategoryAxis.Title =
                excelChart.PrimaryCategoryAxis.Title;
        }
    }
}
