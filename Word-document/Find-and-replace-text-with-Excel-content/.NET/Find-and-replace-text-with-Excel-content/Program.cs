using Syncfusion.XlsIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.DocIORenderer;
using Syncfusion.Drawing;
using Syncfusion.DocIO;


using (FileStream fileStreamPath = new FileStream(Path.GetFullPath(@"Data/Template.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
{
    //Opens an existing Word document.
    using (WordDocument document = new WordDocument(fileStreamPath, FormatType.Docx))
    {
        //Create a new table
        WTable table = new WTable(document);
        //Resizes the table to fit the contents respect to the contents
        table.AutoFit(AutoFitType.FitToContent);

        //Get the Excel content
        ExtractExcelContent(table);

        //Replaces the table placeholder text with a new table.
        TextBodyPart bodyPart = new TextBodyPart(document);
        bodyPart.BodyItems.Add(table);
        document.Replace("<<ExcelPlaceHolder>>", bodyPart, true, true, false);

        //Load the file into stream
        FileStream outputStream = new FileStream(Path.GetFullPath(@"Output/Output.docx"), FileMode.Create, FileAccess.Write);
        //Save the Word document.
        document.Save(outputStream, FormatType.Docx);
    }
}

///<summary>
///Extract the Excel content to Word document table
///</summary>
void ExtractExcelContent(WTable table)
{
    //Open the Excel file
    ExcelEngine excelEngine = new ExcelEngine();
    IApplication application = excelEngine.Excel;
    application.DefaultVersion = ExcelVersion.Xlsx;
    //Load the file into stream
    FileStream inputExcelStream = new FileStream(Path.GetFullPath(@"Data/Sample.xlsx"), FileMode.Open, FileAccess.Read);
    IWorkbook workbook = application.Workbooks.Open(inputExcelStream);

    //Get the first worksheet
    IWorksheet worksheet = workbook.Worksheets[0];
    //Get the number of rows used.
    int rows = worksheet.Rows.Length;
    //Get the number of columns used.
    int columns = worksheet.Columns.Length;

    //Create the rows and columns based on the excel values.
    table.ResetCells(rows, columns);
    //Set the border style
    table.TableFormat.Borders.BorderType = BorderStyle.Single;
    table.TableFormat.Borders.LineWidth = 1;
    table.TableFormat.Borders.Color = Color.Black;

    //Iterate through the excel rows
    for (int rowIndex = 0; rowIndex < rows; rowIndex++)
    {
        //Get the row range from excel
        IRange rowRange = worksheet.Rows[rowIndex];

        //Iterate through the excel columns
        for (int cellIndex = 0; cellIndex < columns; cellIndex++)
        {
            //Get the cell from the excel
            IRange cell = rowRange.Cells[cellIndex];

            //Add a paragraph
            WParagraph paragraph = table[rowIndex, cellIndex].AddParagraph() as WParagraph;
            //Get the content of the cell
            WTextRange textRange = paragraph.AppendText(cell.DisplayText) as WTextRange;
            //Set the bold and italic format
            textRange.CharacterFormat.Bold = cell.CellStyle.Font.Bold;
            textRange.CharacterFormat.Italic = cell.CellStyle.Font.Italic;
            //Get the font size
            textRange.CharacterFormat.FontSize = (float)cell.CellStyle.Font.Size;
            //Get the font name
            textRange.CharacterFormat.FontName = cell.CellStyle.Font.FontName;
        }
    }
}