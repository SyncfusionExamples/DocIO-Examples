using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.Collections;
using System.Data.OleDb;
using System.IO;

namespace Mail_merge_for_nested_group
{
    class Program
    {
        static void Main(string[] args)
        {
            //Opens the template document
			WordDocument document = new WordDocument(Path.GetFullPath(@"../../Data/Template.docx"));
            //Gets the data from the database
            string dataBase = Path.GetFullPath(@"../../Data/EmployeeDetails.mdb");
            OleDbConnection conn = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + dataBase);
			conn.Open();
			//ArrayList contains the list of commands
			ArrayList commands = GetCommands();
			//Executes the mail merge
			document.MailMerge.ExecuteNestedGroup(conn, commands);
			//Saves and closes the Word document instance
			document.Save(Path.GetFullPath(@"../../Result.docx"));
			document.Close();                
        }
        #region Helper Methods
        /// <summary>
        /// Get the commands to execute with database.
        /// </summary>
        private static ArrayList GetCommands()
        {
            //ArrayList contains the list of commands
            ArrayList commands = new ArrayList();
            //DictionaryEntry contains "Source table" (key) and "Command" (value)
            DictionaryEntry entry = new DictionaryEntry("Employees", "SELECT TOP 10 * FROM Employees");
            commands.Add(entry);
            //Retrieves the customer details
            entry = new DictionaryEntry("Customers", "SELECT * FROM Customers WHERE Customers.EmployeeID='%Employees.EmployeeID%'");
            commands.Add(entry);
            //Retrieves the order details
            entry = new DictionaryEntry("Orders", "SELECT * FROM Orders WHERE Orders.CustomerID='%Customers.CustomerID%'");
            commands.Add(entry);
            return commands;
        }
        #endregion

    }
}