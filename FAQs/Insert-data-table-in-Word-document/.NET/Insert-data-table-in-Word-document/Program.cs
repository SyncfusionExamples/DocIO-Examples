using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.Data;
using System.IO;

namespace Insert_data_table_in_Word_document
{
    class Program
    {
        static void Main(string[] args)
        {
            //Creates a new Word document.
            using (WordDocument document = new WordDocument())
            {
                //Creates new data set and data table.
                DataSet dataset = new DataSet();
                GetDataTable(dataset);
                DataTable datatable = new DataTable();
                datatable = dataset.Tables[0];
                //Adds new section.
                IWSection section = document.AddSection();
                //Adds new table.
                IWTable table = section.AddTable();
                //Adds new row to the table.
                WTableRow row = table.AddRow();
                foreach (DataColumn datacolumn in datatable.Columns)
                {
                    //Sets the column names for the table from the data table column names and cell width.
                    WTableCell cell = row.AddCell();
                    cell.AddParagraph().AppendText(datacolumn.ColumnName);
                    cell.Width = 150;
                }
                //Iterates through data table rows.
                foreach (DataRow datarow in datatable.Rows)
                {
                    //Adds new row to the table.
                    row = table.AddRow(true, false);
                    foreach (object datacolumn in datarow.ItemArray)
                    {
                        //Adds new cell.
                        WTableCell cell = row.AddCell();
                        //Adds contents from the data table to the table cell.
                        cell.AddParagraph().AppendText(datacolumn.ToString());
                    }
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
        /// Gets the data table.
        /// </summary>
        /// <param name="dataset"></param>
        private static void GetDataTable(DataSet dataset)
        {
            //List of syncfusion products.
            string[] products = { "DocIO", "PDF", "XlsIO" };
            //Adds new Tables to the data set.
            DataRow row;
            dataset.Tables.Add();
            //Adds fields to the Products table.
            dataset.Tables[0].TableName = "Products";
            dataset.Tables[0].Columns.Add("ProductName");
            dataset.Tables[0].Columns.Add("Binary");
            dataset.Tables[0].Columns.Add("Source");
            //Inserts values to the tables.
            foreach (string product in products)
            {
                row = dataset.Tables["Products"].NewRow();
                row["ProductName"] = string.Concat("Essential ", product);
                row["Binary"] = "$895.00";
                row["Source"] = "$1,295.00";
                dataset.Tables["Products"].Rows.Add(row);
            }
        }
    }
}
