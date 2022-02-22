using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.Data;
using System.IO;

namespace Open_and_save_macro_enabled_document
{
    class Program
    {
        static void Main(string[] args)
        {
            using (FileStream fileStream = new FileStream(Path.GetFullPath(@"../../../Template.dotm"), FileMode.Open, FileAccess.ReadWrite))
            {
                //Opens the template document.
                using (WordDocument document = new WordDocument(fileStream, FormatType.Dotm))
                {
                    //Gets the table.
                    DataTable table = GetDataTable();
                    //Executes Mail Merge with groups.
                    document.MailMerge.ExecuteGroup(table);
                    //Creates file stream.
                    using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"../../../Result.docm"), FileMode.Create, FileAccess.ReadWrite))
                    {
                        //Saves the Word document to file stream.
                        document.Save(outputFileStream, FormatType.Word2013Docm);
                    }
                }
            }
        }

        private static DataTable GetDataTable()
        {
            //List of syncfusion products name.
            string[] products = { "DocIO", "PDF", "XlsIO" };
            //Adds new Tables to the data set.
            DataRow row;
            DataTable table = new DataTable();
            //Adds fields to the Products table.
            table.TableName = "Products";
            table.Columns.Add("ProductName");
            table.Columns.Add("Binary");
            table.Columns.Add("Source");
            //Inserts values to the tables.
            foreach (string product in products)
            {
                row = table.NewRow();
                row["ProductName"] = string.Concat("Essential ", product);
                row["Binary"] = "$895.00";
                row["Source"] = "$1,295.00";
                table.Rows.Add(row);
            }
            return table;
        }
    }
}
