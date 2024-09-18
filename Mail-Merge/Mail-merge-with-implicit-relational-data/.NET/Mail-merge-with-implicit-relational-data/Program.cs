using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.Collections.Generic;
using System.IO;

namespace Mail_merge_with_implicit_relational_data
{
    class Program
    {
        static void Main(string[] args)
        {
            using (FileStream fileStream = new FileStream(Path.GetFullPath(@"Data/Template.docx"), FileMode.Open, FileAccess.ReadWrite))
            {
                //Opens the template document.
                using (WordDocument document = new WordDocument(fileStream, FormatType.Docx))
                {
                    //Gets the organization details as “IEnumerable” collection.
                    List<Organization> organizationList = GetOrganizations();
                    //Creates an instance of “MailMergeDataTable” by specifying mail merge group name and “IEnumerable” collection.
                    MailMergeDataTable dataTable = new MailMergeDataTable("Organizations", organizationList);
                    //Performs Mail merge.
                    document.MailMerge.ExecuteNestedGroup(dataTable);
                    //Creates file stream.
                    using (FileStream outputStream = new FileStream(Path.GetFullPath(@"Output/Result.docx"), FileMode.Create, FileAccess.ReadWrite))
                    {
                        //Saves the Word document to file stream.
                        document.Save(outputStream, FormatType.Docx);
                    }
                }
            }
        }

        #region Helper Methods
        /// <summary>
        /// Get the data to perform mail merge.
        /// </summary>
        public static List<Organization> GetOrganizations()
        {
            //Creates Employee details.
            List<EmployeeDetails> employees = new List<EmployeeDetails>();
            employees.Add(new EmployeeDetails("Thomas Hardy", "1001", "05/27/1996"));
            employees.Add(new EmployeeDetails("Maria Anders", "1002", "04/10/1998"));
            //Creates Departments details.
            List<DepartmentDetails> departments = new List<DepartmentDetails>();
            departments.Add(new DepartmentDetails("Marketing", "Nancy Davolio", employees));

            employees = new List<EmployeeDetails>();
            employees.Add(new EmployeeDetails("Elizabeth Lincoln", "1003", "05/15/1996"));
            employees.Add(new EmployeeDetails("Antonio Moreno", "1004", "04/22/1996"));
            departments.Add(new DepartmentDetails("Production", "Andrew Fuller", employees));
            //Creates organization details.
            List<Organization> organizations = new List<Organization>();
            organizations.Add(new Organization("UK Office", "120 Hanover Sq.", "London", "WX1 6LT", "UK", departments));
            return organizations;
        }
        #endregion
    }

    #region Helper class
    /// <summary>
    /// Represents a class to maintain organization details.
    /// </summary>
    public class Organization
    {
        public string BranchName { get; set; }
        public string Address { get; set; }
        public string City { get; set; }
        public string ZipCode { get; set; }
        public string Country { get; set; }
        public List<DepartmentDetails> Departments { get; set; }
        public Organization(string branchName, string address, string city, string zipcode, string country, List<DepartmentDetails> departments)
        {
            BranchName = branchName;
            Address = address;
            City = city;
            ZipCode = zipcode;
            Country = country;
            Departments = departments;
        }
    }
    /// <summary>
    /// Represents a class to maintain department details.
    /// </summary>
    public class DepartmentDetails
    {
        public string DepartmentName { get; set; }
        public string Supervisor { get; set; }
        public List<EmployeeDetails> Employees { get; set; }
        public DepartmentDetails(string departmentName, string supervisor, List<EmployeeDetails> employees)
        {
            DepartmentName = departmentName;
            Supervisor = supervisor;
            Employees = employees;
        }
    }
    /// <summary>
    /// Represents a class to maintain employee details.
    /// </summary>
    public class EmployeeDetails
    {
        public string EmployeeName { get; set; }
        public string EmployeeID { get; set; }
        public string JoinedDate { get; set; }
        public EmployeeDetails(string employeeName, string employeeID, string joinedDate)
        {
            EmployeeName = employeeName;
            EmployeeID = employeeID;
            JoinedDate = joinedDate;
        }
    }
    #endregion
}
