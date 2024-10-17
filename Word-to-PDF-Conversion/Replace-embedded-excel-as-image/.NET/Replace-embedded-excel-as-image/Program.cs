using Syncfusion.Pdf;
using Syncfusion.DocIORenderer;
using Syncfusion.DocIO.DLS;
using Syncfusion.DocIO;
using Syncfusion.XlsIO;
using Syncfusion.XlsIORenderer;

//Initialize the DocIORenderer component for converting Word documents to PDF.
using (DocIORenderer docIORenderer = new DocIORenderer())
{
    using (FileStream inputStream = new FileStream(Path.GetFullPath(@"Data/Input.docx"), FileMode.Open, FileAccess.Read))
    {
        //Open the input Word document.
        using (WordDocument document = new WordDocument(inputStream, FormatType.Automatic))
        {
            //Replace embedded Excel objects in the document with images.
            ReplaceExcelToImage(document);
            //Convert the Word document to a PDF using the DocIORenderer component.
            using (PdfDocument pdf = docIORenderer.ConvertToPDF(document))
            {
                using (FileStream outputStream = new FileStream(Path.GetFullPath(@"Output/Output.pdf"), FileMode.Create, FileAccess.Write))
                {
                    //Save the generated PDF to the specified file stream.
                    pdf.Save(outputStream);
                }
            }
        }
    }
}


/// <summary>
/// Replaces embedded Excel OLE objects in a Word document with their corresponding images while maintaining the original size.
/// </summary>
void ReplaceExcelToImage(WordDocument wordDocument)
{
    //Get the Ole objects.
    List<Entity> oleObjects = wordDocument.FindAllItemsByProperty(EntityType.OleObject, null, null);
    //Iterate through the ole objects.
    for (int i = 0; i < oleObjects.Count; i++)
    {
        WOleObject ole = oleObjects[i] as WOleObject;
        //Check the type of OLE.
        string type = ole.ObjectType;
        //Get the height and width of OLE picture.
        float height = ole.OlePicture.Height;
        float width = ole.OlePicture.Width;
        //If the type contains "Excel", then the OLE object is extracted from Excel.
        if (type.Contains("Excel"))
        {
            //Create a Excel file using the Ole data.
            MemoryStream excelStream = new MemoryStream();
            excelStream.Write(ole.NativeData);
            excelStream.Position = 0;

            //Creates a new instance for ExcelEngine.
            ExcelEngine excelEngine = new ExcelEngine();
            //Initialize IApplication.
            IApplication application = excelEngine.Excel;
            //Loads or open an existing workbook through Open method of IWorkbooks.
            IWorkbook workbook = application.Workbooks.Open(excelStream);
            IWorksheet sheet = workbook.Worksheets[0];

            //Initialize XlsIORenderer.
            application.XlsIORenderer = new XlsIORenderer();

            //Converts and save as stream.
            MemoryStream imgStream = new MemoryStream();
            sheet.ConvertToImage(1, 1, 6, 5, imgStream);
            imgStream.Position = 0;

            //Load the converted image as OLE picture.
            ole.OlePicture.LoadImage(imgStream);
            ole.OlePicture.LockAspectRatio = false;
            ole.OlePicture.Height = height;
            ole.OlePicture.Width = width;

            //Close and Dispose.
            workbook.Close();
            imgStream.Dispose();
            excelStream.Dispose();
        }
    }
}
