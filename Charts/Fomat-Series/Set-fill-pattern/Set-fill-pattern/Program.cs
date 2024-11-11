using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.OfficeChart;

using (FileStream fileStreamPath = new FileStream(@"Data/Template.docx", FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
{
    //Open the template Word document
    using (WordDocument document = new WordDocument(fileStreamPath, FormatType.Docx))
    {
        //Gets the last paragraph
        WParagraph paragraph = document.LastParagraph;
        //Gets the chart entity from the paragraph items
        WChart chart = paragraph.ChildEntities[0] as WChart;

        // Set the fill pattern for the series in the chart.
        chart.Series[0].SerieFormat.Fill.Pattern = OfficeGradientPattern.Pat_Diagonal_Brick;
        chart.Series[1].SerieFormat.Fill.Pattern = OfficeGradientPattern.Pat_Dashed_Vertical;
        chart.Series[2].SerieFormat.Fill.Pattern = OfficeGradientPattern.Pat_Sphere;

        using (FileStream stream = new FileStream(Path.GetFullPath(@"Output/Result.docx"), FileMode.Create, FileAccess.Write))
        {
            //Save the Word document.
            document.Save(stream, FormatType.Docx);
        }
    }

}
