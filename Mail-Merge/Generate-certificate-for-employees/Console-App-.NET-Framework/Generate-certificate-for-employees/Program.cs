using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.Data.OleDb;
using System.IO;

namespace Generate_certificates_for_employees
{
    class Program
    {
        static void Main(string[] args)
        {
            //Creates new Word document instance for Word processing
            using (WordDocument document = new WordDocument())
            {
                //Opens the Word template document
                document.Open(Path.GetFullPath(@"../../Template.docx"), FormatType.Docx);
                //Loads the database
                OleDbConnection conn = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + @"../../EmployeeList.mdb");
                //Opens the database connection
                conn.Open();
                OleDbCommand command = new OleDbCommand("Select * from Employees", conn);
                //Executes the command to read the data from database
                OleDbDataReader reader = command.ExecuteReader();
                //Perform mail merge
                document.MailMerge.Execute(reader);
                //Dispose the command
                command.Dispose();
                //Closes the reader
                reader.Close();
                //Closes the database connection
                conn.Dispose();
                //Saves the Word document
                document.Save(Path.GetFullPath(@"../../Sample.docx"));
            }
        }
        
    }
}
