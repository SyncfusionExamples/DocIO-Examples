using Syncfusion.DocIO.DLS;
using Syncfusion.XlsIO;
using System.Data;
using System.IO;
using System.Reflection;

namespace MailMergeWithExcelData
{
    class Program
    {
        static void Main(string[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                //Instantiate the Excel application object
                IApplication application = excelEngine.Excel;

                //Load an existing Excel file into IWorkbook
                Assembly assembly = typeof(Program).GetTypeInfo().Assembly;
                Stream excelStream = assembly.GetManifestResourceStream("ExportExcelToWord.NorthwindDataTemplate.xls");
                IWorkbook workbook = application.Workbooks.Open(excelStream, ExcelOpenType.Automatic);

                //Get the first worksheet in workbook into IWorksheet
                IWorksheet worksheet = workbook.Worksheets[0];

                //Initialize the DataTable
                DataTable dataTable = new DataTable();

                //Export the data in Excel worksheet into DataTable and assign table name
                dataTable = worksheet.ExportDataTable(worksheet.UsedRange, ExcelExportDataTableOptions.ColumnNames);
                dataTable.TableName = "Customers";

                //Load an existing Word document into WordDocument
                Stream wordStream = assembly.GetManifestResourceStream("ExportExcelToWord.CustomersReport.doc");
                WordDocument wordDocument = new WordDocument(wordStream);

                //Export the data in DataTable into Word document through MailMerge
                wordDocument.MailMerge.Execute(dataTable);

                //Save the Word document and close the instance of WrodDocument
                wordDocument.Save("Output.docx");
                wordDocument.Close();

                System.Diagnostics.Process.Start("Output.docx");
            }
        }
    }
}