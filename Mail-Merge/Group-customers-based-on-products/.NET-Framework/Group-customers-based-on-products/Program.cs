using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.Collections;
using System.Data;
using System.IO;

namespace Group_customers_based_on_products
{
    class Program
    {
        static void Main(string[] args)
        {
            //Creates new Word document instance for Word processing.
            using (WordDocument document = new WordDocument())
            {
                //Opens the Word template document.
                document.Open(Path.GetFullPath(@"../../Template.docx"), FormatType.Docx);
                //Creates an instance for DataSet.
                DataSet dataSet = new DataSet();
                //Get the customer details.
                DataTable customerData = GetCustomerDetails();
                //Adds the DataTable in DataSet.
                dataSet.Tables.Add(customerData);
                //Gets the product details.
                DataTable productData = GetProductDetails();
                //Adds the DataTable in DataSet.
                dataSet.Tables.Add(productData);                
                //ArrayList contains the list of commands.
                ArrayList commands = new ArrayList();
                //DictionaryEntry contain "Source table" (KEY) and "Command" (VALUE).
                DictionaryEntry entry = new DictionaryEntry("ProductList", string.Empty);
                commands.Add(entry);
                //To retrive customer details that match the Product.
                entry = new DictionaryEntry("Customers", "ProductName = %ProductList.ProductName%");
                commands.Add(entry);
                //Perform nested mail merge.
                document.MailMerge.ExecuteNestedGroup(dataSet, commands);
                //Saves the Word document.
                document.Save(Path.GetFullPath(@"../../Sample.docx"), FormatType.Docx);

            }
        }
        #region Helper methods
        /// <summary>
        /// Gets the data to perform mail merge.
        /// </summary>
        /// <returns></returns>
        private static DataTable GetCustomerDetails()
        {
            //Creates new DataTable instance. 
            DataTable table = new DataTable("Customers");
            //Add columns for the DataTable.
            table.Columns.Add("CustomerId");
            table.Columns.Add("CustomerName");
            table.Columns.Add("ProductName");

            //Add records in DataTable.
            DataRow row = table.NewRow();
            row["CustomerId"] = "1001";
            row["CustomerName"] = "Diego Roel";
            row["ProductName"] = "Essential DocIO";
            table.Rows.Add(row);

            row = table.NewRow();
            row["CustomerId"] = "1002";
            row["CustomerName"] = "Maria Larsson";
            row["ProductName"] = "Essential DocIO";
            table.Rows.Add(row);

            row = table.NewRow();
            row["CustomerId"] = "1003";
            row["CustomerName"] = "Pedro Afonso";
            row["ProductName"] = "Essential XlsIO";
            table.Rows.Add(row);

            row = table.NewRow();
            row["CustomerId"] = "1004";
            row["CustomerName"] = "Maria Larsson";
            row["ProductName"] = "Essential XlsIO";
            table.Rows.Add(row);

            row = table.NewRow();
            row["CustomerId"] = "1005";
            row["CustomerName"] = "Paolo Accorti";
            row["ProductName"] = "Essential DocIO";
            table.Rows.Add(row);

            row = table.NewRow();
            row["CustomerId"] = "1006";
            row["CustomerName"] = "Mario Pontes";
            row["ProductName"] = "Essential XlsIO";
            table.Rows.Add(row);

            row = table.NewRow();
            row["CustomerId"] = "1007";
            row["CustomerName"] = "Daniel Tonini";
            row["ProductName"] = "Essential PDF";
            table.Rows.Add(row);

            row = table.NewRow();
            row["CustomerId"] = "1008";
            row["CustomerName"] = "Felipe Izquierdo";
            row["ProductName"] = "Essential PDF";
            table.Rows.Add(row);

            row = table.NewRow();
            row["CustomerId"] = "1009";
            row["CustomerName"] = "Bernardo Batista";
            row["ProductName"] = "Essential PDF";
            table.Rows.Add(row);

            row = table.NewRow();
            row["CustomerId"] = "1010";
            row["CustomerName"] = "Maurizio Moroni";
            row["ProductName"] = "Essential DocIO";
            table.Rows.Add(row);

            return table;
        }

        /// <summary>
        /// Gets the data to perform mail merge.
        /// </summary>
        /// <returns></returns>
        private static DataTable GetProductDetails()
        {
            //Creates new DataTable instance. 
            DataTable table = new DataTable("ProductList");
            //Add columns in DataTable.
            table.Columns.Add("ProductName");
            //Add records in DataTable.
            DataRow row = table.NewRow();
            row["ProductName"] = "Essential DocIO";
            table.Rows.Add(row);

            row = table.NewRow();
            row["ProductName"] = "Essential XlsIO";
            table.Rows.Add(row);

            row = table.NewRow();
            row["ProductName"] = "Essential PDF";
            table.Rows.Add(row);
            return table;
        }
        
        #endregion
    }
}
