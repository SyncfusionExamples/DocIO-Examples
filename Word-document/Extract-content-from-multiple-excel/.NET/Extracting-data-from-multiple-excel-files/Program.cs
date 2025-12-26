using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.DocIORenderer;
using Syncfusion.Drawing;
using Syncfusion.XlsIO;


namespace Extract_data_from_multiple_excel
{
    class Program
    {
        static void Main(string[] args)
        {
            //Creates a new Word document.
            using (WordDocument document = new WordDocument())
            {
                // List of Excel file names (without extension)
                string[] excelFiles = { "Excel1", "Excel2" };
                // Loop through each Excel file
                foreach (string excelName in excelFiles)
                {
                    // Get the Excel content
                    UpdateExcelToWord(document, excelName);
                }
                //Creates file stream.
                using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"Output/Output.docx"), FileMode.Create, FileAccess.ReadWrite))
                {
                    //Saves the Word document to file stream.
                    document.Save(outputFileStream, FormatType.Docx);
                }
            }

        }

        /// <summary>
        /// Reads an Excel workbook and appends each worksheet as a table into the specified Word document.
        /// </summary>
        /// <param name="document"></param>

        private static void UpdateExcelToWord(WordDocument document, string excelName)
        {
            // Open the Excel file
            ExcelEngine excelEngine = new ExcelEngine();
            IApplication application = excelEngine.Excel;
            application.DefaultVersion = ExcelVersion.Xlsx;
            // Load the file into stream
            using (FileStream inputExcelStream = new FileStream(Path.GetFullPath(@"Data/") + excelName + ".xlsx", FileMode.Open, FileAccess.Read))
            {
                // Open the workbook from the stream.
                IWorkbook workbook = application.Workbooks.Open(inputExcelStream);
                // Loop through all worksheets
                foreach (IWorksheet worksheet in workbook.Worksheets)
                {
                    // Add a new section for each worksheet
                    IWSection section = document.AddSection();
                    // Add a new table to the section
                    WTable table = section.AddTable() as WTable;
                    table.AutoFit(AutoFitType.FitToContent);
                    //Set Title for table
                    table.Title = excelName + "_" + worksheet.Name;
                    // Determine the number of rows and columns in the worksheet.
                    int rows = worksheet.Rows.Length;
                    int columns = worksheet.Columns.Length;
                    // Initialize the table with the required number of rows and columns.
                    table.ResetCells(rows, columns);
                    table.TableFormat.Borders.BorderType = BorderStyle.Single;
                    table.TableFormat.Borders.LineWidth = 1;
                    table.TableFormat.Borders.Color = Color.Black;
                    // Populate the table cell-by-cell from the worksheet.
                    for (int rowIndex = 0; rowIndex < rows; rowIndex++)
                    {
                        // Get the current row range in the worksheet.
                        IRange rowRange = worksheet.Rows[rowIndex];

                        for (int cellIndex = 0; cellIndex < columns; cellIndex++)
                        {
                            // Access the specific cell within the current row.
                            IRange cell = rowRange.Cells[cellIndex];
                            // Add a paragraph into the corresponding Word table cell.
                            WParagraph paragraph = table[rowIndex, cellIndex].AddParagraph() as WParagraph;
                            // Insert the cell's display text into the Word paragraph.
                            WTextRange textRange = paragraph.AppendText(cell.DisplayText) as WTextRange;
                            // Preserve basic font styling from the Excel cell into the Word text.
                            textRange.CharacterFormat.Bold = cell.CellStyle.Font.Bold;
                            textRange.CharacterFormat.Italic = cell.CellStyle.Font.Italic;
                            textRange.CharacterFormat.FontSize = (float)cell.CellStyle.Font.Size;
                            textRange.CharacterFormat.FontName = cell.CellStyle.Font.FontName;
                        }
                    }
                }
            }              
        }
    }
}

