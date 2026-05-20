using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.Collections;
using System.Data;
using System.Data.OleDb;

namespace Convert_Database_into_Dataset_and_Perform_Mailmerge
{
    class Program
    {
        static void Main(string[] args)
        {
            // Load the word document
            using (FileStream sourceStreamPath = new FileStream(Path.GetFullPath(@"../../../Data/Template.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                //Creating a new document.
                using (WordDocument document = new WordDocument(sourceStreamPath, FormatType.Docx))
                {

                    string dataBase = Path.GetFullPath(@"../../../Data/EmployeeDetails.mdb");

                    // Get all data
                    DataSet ds = GetAllTables(dataBase);
                    //ArrayList contains the list of commands
                    ArrayList commands = GetCommands();
                    //Executes the mail merge
                    document.MailMerge.ExecuteNestedGroup(ds, commands);
                    //Saves the Word document to MemoryStream.
                    using (FileStream outputStream = new FileStream(Path.GetFullPath(@"../../../Output/Result.docx"), FileMode.Create, FileAccess.ReadWrite))
                    {
                        document.Save(outputStream, FormatType.Docx);
                    }
                }
            }
        } 

        /// <summary>
        /// Get the commands to execute with database.
        /// </summary>
        static ArrayList GetCommands()
        {
            //ArrayList contains the list of commands
            ArrayList commands = new ArrayList();

            // Parent table: Employees (no filter, so empty string)
            commands.Add(new DictionaryEntry("Employees", ""));

            // Customers filtered by EmployeeID
            commands.Add(new DictionaryEntry("Customers", "EmployeeID = %Employees.EmployeeID%"));

            // Orders filtered by CustomerID
            commands.Add(new DictionaryEntry("Orders", "CustomerID = %Customers.CustomerID%"));

            return commands;
        }
        //Retrieves all required tables from the MDB database and prepares hierarchy commands for DocIO mail merge.
        static DataSet GetAllTables(string dataBase)
        {
            // Connection string using ACE OLEDB provider (64-bit compatible)
            string connectionString = $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={dataBase};Persist Security Info=False;";
            // DataSet to hold all tables
            DataSet ds = new DataSet();

            using (OleDbConnection conn = new OleDbConnection(connectionString))
            {
                conn.Open();
                // Tables to fetch from MDB
                string[] tables = { "Employees", "Customers", "Orders" };
                // Loop through each table and fill DataSet
                foreach (string tableName in tables)
                {
                    string sqlQuery = $"SELECT * FROM {tableName}";
                    OleDbDataAdapter adapter = new OleDbDataAdapter(sqlQuery, conn);
                    adapter.Fill(ds, tableName);
                }
            }
            return ds;
        }
    }
}
