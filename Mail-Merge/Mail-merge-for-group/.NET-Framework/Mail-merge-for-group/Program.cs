using Syncfusion.DocIO.DLS;
using System.Data;
using System.Data.SqlServerCe;
using System.IO;

namespace Mail_merge_for_group
{
    class Program
    {
        static void Main(string[] args)
        {
            //Opens the template document
            WordDocument document = new WordDocument(Path.GetFullPath(@"../../Data/EmployeesTemplate.docx"));
            //Gets the data table
            DataTable table = GetDataTable();
            //Executes Mail Merge with groups
            document.MailMerge.ExecuteGroup(table);
            //Saves and closes the WordDocument instance
            document.Save(Path.GetFullPath(@"../../Result.docx"));
            document.Close();
        }

        #region Helper methods
        /// <summary>
        /// Gets the data to perform mail merge.
        /// </summary>
        /// <returns></returns>
        private static DataTable GetDataTable()
        {
            string datasourceName = Path.GetFullPath(@"../../Data/EmployeeDetails.sdf");
            DataSet dataset = new DataSet();
            SqlCeConnection conn = new SqlCeConnection("Data Source = " + datasourceName);
            conn.Open();
            SqlCeDataAdapter adapter = new SqlCeDataAdapter("Select TOP(5) * from EmployeesReport", conn);
            adapter.Fill(dataset);
            adapter.Dispose();
            conn.Close();
            DataTable table = dataset.Tables[0];
            //Sets table name as Employees for template merge field reference
            table.TableName = "Employees";
            return table;
        }
        #endregion
    }
}