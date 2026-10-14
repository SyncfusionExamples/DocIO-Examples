using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.XlsIO;
using System.Collections;
using System.Data;
using System.IO;

namespace Perform_Mail_Merge_With_Excel_Data
{
    class Program
    {
        static void Main(string[] args)
        {

            // Creates a new instance of FileStream to open the template document.
            using (FileStream fileStream = new FileStream(Path.GetFullPath(@"Data/Template.docx"), FileMode.Open, FileAccess.ReadWrite))
            {
                //Opens the template document.
                using (WordDocument document = new WordDocument(fileStream, FormatType.Docx))
                {
                    //Uses the mail merge events handler for image fields.
                    document.MailMerge.MergeImageField += new MergeImageFieldEventHandler(MergeField_LogoImage);
                    document.MailMerge.StartAtNewPage = true;
                    //Get the data table from the Excel file
                    DataTable dataTable = GetDataTable(Path.GetFullPath(@"Data/InvoiceGroupDetails.xlsx"));
                    //Set the table name for the data table.
                    //Note that table name and group name should be same.
                    dataTable.TableName = "Invoice";
                    //Create data set and add data table into it
                    DataSet dataSet = new DataSet();
                    dataSet.Tables.Add(dataTable);
                    //Get the data table from the Excel file
                    dataTable = GetDataTable(Path.GetFullPath(@"Data/ProductDetails.xlsx"));
                    dataTable.TableName = "Products";
                    dataSet.Tables.Add(dataTable);
                    //Get the commands for mail merge execution.
                    ArrayList commands = GetCommands();
                    //Executes the nested mail merge with the specified data set and commands.
                    document.MailMerge.ExecuteNestedGroup(dataSet, commands);
                    document.Save(Path.GetFullPath(@"Output/Invoice.docx"), FormatType.Docx);
                }
            }
        }

        #region Helper Methods
        /// <summary>
        /// Read the Excel file and return the data table for Mail Merge.
        /// </summary>
        /// <param name="fileName"></param>
        /// <returns></returns>
        private static DataTable GetDataTable(string fileName)
        {
            DataTable dataTable;
            //Creates a new instance for FileStream to read the Excel file.
            using (FileStream fileStream = new FileStream(fileName, FileMode.Open, FileAccess.ReadWrite))
            {
                //Creates a new instance for ExcelEngine.
                using (ExcelEngine excelEngine = new ExcelEngine())
                {
                    //Loads or open an existing workbook through Open method of IWorkbooks.
                    IWorkbook workbook = excelEngine.Excel.Workbooks.Open(fileStream);
                    //The first worksheet object in the worksheets collection is accessed.
                    IWorksheet sheet = workbook.Worksheets[0];
                    //Get as DataTable.
                    dataTable = sheet.ExportDataTable(sheet.UsedRange, ExcelExportDataTableOptions.ColumnNames);
                }
            }
            return dataTable;
        }
        /// <summary>
        /// Get the commands for Mail Merge Execution.
        /// </summary>
        /// <returns></returns>
        private static ArrayList GetCommands()
        {
            //Define commands with the table name and expression for linking the multiple data tables
            //during nested Mail merge process.
            //You can use the “%TableName.ColumnName%” expression for getting the current value of specified column or field from the table.

            //ArrayList contains the list of commands
            ArrayList commands = new ArrayList();
            //DictionaryEntry contains "Source table" (key) and "Command" (value)
            //Retrieves the invoice details
            DictionaryEntry entry = new DictionaryEntry("Invoice", string.Empty);
            commands.Add(entry);
            //Retrieves the products details
            entry = new DictionaryEntry("Products", "InvoiceNo = %Invoice.InvoiceNo%");
            commands.Add(entry);
            return commands;
        }
        /// <summary>
        /// Mail merge events handler for image fields.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="args"></param>
        private static void MergeField_LogoImage(object sender, MergeImageFieldEventArgs args)
        {
            //Binds image from file system during mail merge.
            if (args.FieldName == "Logo")
            {
                string logoFileName = args.FieldValue.ToString();
                //Gets the image from file system
                FileStream imageStream = new FileStream(@"../../../Data/" + logoFileName, FileMode.Open, FileAccess.Read);
                args.ImageStream = imageStream;
                //Gets the picture, to be merged for image merge field.
                WPicture picture = args.Picture;
                //Resizes the picture
                picture.Height = 40;
                picture.Width = 90;
            }
        }
        #endregion

    }
}
