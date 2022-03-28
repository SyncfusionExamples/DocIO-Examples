using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.Data;
using System.Data.OleDb;
using System.IO;

namespace Mail_merge_using_OleDbConnection
{
    class Program
    {
        static void Main(string[] args)
        {
            string dataBase = Path.GetFullPath(@"../../Data/EmployeeDetails.mdb");
            //Opens existing template.
            using (WordDocument document = new WordDocument(Path.GetFullPath(@"../../Data/Template.docx"), FormatType.Docx))
            {
                //Gets Data from the Database.
                OleDbConnection conn = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + dataBase);
                conn.Open();
                //Populates the data table.
                DataTable table = new DataTable();
                OleDbDataAdapter adapter = new OleDbDataAdapter("select * from employees", conn);
                adapter.Fill(table);
                adapter.Dispose();
                //Performs Mail Merge.
                document.MailMerge.Execute(table);
                //Saves and closes the document.
                document.Save(Path.GetFullPath(@"../../Result.docx"), FormatType.Docx);
            }
        }
    }
}
