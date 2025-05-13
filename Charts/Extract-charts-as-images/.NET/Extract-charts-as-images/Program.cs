using Syncfusion.DocIO.DLS;
using Syncfusion.DocIORenderer;

// Load the Word document.
using (WordDocument document = new WordDocument(Path.GetFullPath(@"Data/Template.docx")))
{
    // Find all chart elements in the Word document by EntityType.
    List<Entity> charts = document.FindAllItemsByProperty(EntityType.Chart, null, null);

    // Create an instance of DocIORenderer.
    using (DocIORenderer renderer = new DocIORenderer())
    {
        // Loop through each chart found in the document.
        for (int i = 0; i < charts.Count; i++)
        {
            WChart chart = charts[i] as WChart;

            // Convert the chart to an image stream.
            using (Stream stream = chart.SaveAsImage())
            {
                // Create a file stream to save the image to disk.
                using (FileStream fileStreamOutput = File.Create(Path.GetFullPath(@"Output/Chart" + i + ".jpeg")))
                {
                    // Copy the image stream to the output file.
                    stream.CopyTo(fileStreamOutput);
                }
            }
        }
    }
}
