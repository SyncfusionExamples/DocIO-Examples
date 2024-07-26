using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.Collections.Generic;
using System.IO;

namespace Restart_list_numbering_in_mail_merge
{
    class Program
    {
        static void Main(string[] args)
        {
            using (FileStream fileStream = new FileStream(Path.GetFullPath(@"../../../Template.docx"), FileMode.Open, FileAccess.ReadWrite))
            {
                //Opens the template document.
                using (WordDocument document = new WordDocument(fileStream, FormatType.Docx))
                {
                    //Sets ImportOptions to restart the list numbering.
                    document.ImportOptions = ImportOptions.ListRestartNumbering;
                    //Creates the employee details as “IEnumerable” collection.
                    List<Employee> employeeList = new List<Employee>();
                    employeeList.Add(new Employee("101", "Nancy Davolio", "Seattle, WA, USA"));
                    employeeList.Add(new Employee("102", "Andrew Fuller", "Tacoma, WA, USA"));
                    employeeList.Add(new Employee("103", "Janet Leverling", "Kirkland, WA, USA"));
                    //Creates an instance of “MailMergeDataTable” by specifying mail merge group name and “IEnumerable” collection.
                    MailMergeDataTable dataTable = new MailMergeDataTable("Employee", employeeList);
                    //Performs mail merge.
                    document.MailMerge.ExecuteGroup(dataTable);
                    //Creates file stream.
                    using (FileStream outputStream = new FileStream(Path.GetFullPath(@"../../../Result.docx"), FileMode.Create, FileAccess.ReadWrite))
                    {
                        //Saves the Word document to file stream.
                        document.Save(outputStream, FormatType.Docx);
                    }
                }
            }
        }        
    }

    #region Helper Class
    /// <summary>
    /// Represents the helper class to perform mail merge.
    /// </summary>
    public class Employee
    {
        public string EmployeeID { get; set; }
        public string EmployeeName { get; set; }
        public string Location { get; set; }

        /// <summary>
        /// Represents a constructor to create value for merge fields
        /// </summary>  
        public Employee(string employeeId, string employeeName, string location)
        {
            EmployeeID = employeeId;
            EmployeeName = employeeName;
            Location = location;
        }
    }
    #endregion
}
