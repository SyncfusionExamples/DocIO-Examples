using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.Data;
using System.Data.OleDb;
using System.IO;

namespace Generate_payroll_for_employees
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
                OleDbConnection conn = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + @"../../EmployeeDetails.mdb");
                //Opens the database connection
                conn.Open();
                //Creates command to retrieve data from database
                OleDbCommand cmd = new OleDbCommand("SELECT * FROM Employees", conn);
                //Executes the command in database
                IDataReader dataReader = cmd.ExecuteReader();
                //Performs mail merge
                document.MailMerge.Execute(dataReader);
                //Saves the Word document
                document.Save(Path.GetFullPath(@"../../Sample.docx"), FormatType.Docx);
            }
        }      
    }
}
