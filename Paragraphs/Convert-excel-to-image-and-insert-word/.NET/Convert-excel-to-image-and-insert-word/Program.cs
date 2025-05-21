using Syncfusion.XlsIO;
using Syncfusion.XlsIORenderer;
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;


// Initialize XlsIO renderer
using (ExcelEngine excelEngine = new ExcelEngine())
{
    IApplication application = excelEngine.Excel;

    // Initialize the XlsIORenderer (required for ConvertToImage)
    application.XlsIORenderer = new XlsIORenderer();

    // Open the Excel file
    IWorkbook workbook = application.Workbooks.Open(Path.GetFullPath(@"Data/Template.xlsx"));

    // Create the Word document
    using (WordDocument document = new WordDocument())
    {
        for (int i = 0; i < workbook.Worksheets.Count; i++)
        {
            IWorksheet worksheet = workbook.Worksheets[i];

            // Convert used range of worksheet to image
            using (MemoryStream imageStream = new MemoryStream())
            {
                worksheet.ConvertToImage(worksheet.UsedRange, imageStream);
                imageStream.Position = 0;

                // Add a new section for each worksheet
                IWSection section = document.AddSection();
                section.PageSetup.Orientation = PageOrientation.Landscape;
                IWParagraph paragraph = section.AddParagraph();

                // Insert the Excel image into the paragraph
                IWPicture picture = paragraph.AppendPicture(imageStream);

                // Optionally set image dimensions
                picture.Width = 600;
                picture.Height = 380;
            }
        }

        // Save the Word document
        document.Save(Path.GetFullPath(@"Output/Result.docx"), FormatType.Docx);
    }
}
