using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.Collections.Generic;
using System.Data;
using System.IO;

namespace Mail_merge_with_two_data_sources
{
    class Program
    {
        static void Main(string[] args)
        {
            // Load the Word document
            using (WordDocument document = new WordDocument(Path.GetFullPath(@"Data/EmployeesTemplate.docx")))
            {
                //Sets “ClearFields” to true to remove empty mail merge fields from document 
                document.MailMerge.ClearFields = false;
                //Gets the employee details as IEnumerable collection.
                List<Employee> employeeList = GetEmployees();
                //Creates an instance of MailMergeDataTable by specifying MailMerge group name and IEnumerable collection.
                MailMergeDataTable dataSource = new MailMergeDataTable("Employees", employeeList);
                //Performs Mail merge.
                document.MailMerge.ExecuteGroup(dataSource);

                //Uses the mail merge events handler for image fields.
                document.MailMerge.MergeImageField += new MergeImageFieldEventHandler(MergeField_LogoImage);
                //Gets the DataTable
                DataTable dataTable = GetDataTable();
                //Performs mail merge to merge the logo
                document.MailMerge.Execute(dataTable);

                // Save the modified document
                document.Save(Path.GetFullPath(@"../../../Output/Result.docx"), FormatType.Docx);
            }
        }
        /// <summary>
        /// Gets the employee details to perform mail merge.
        /// </summary>
        public static List<Employee> GetEmployees()
        {
            List<Employee> employees = new List<Employee>();
            employees.Add(new Employee("Nancy", "Smith", "Sales Representative", "505 - 20th Ave. E. Apt. 2A,", "Seattle", "WA", "USA"));
            employees.Add(new Employee("Andrew", "Fuller", "Vice President, Sales", "908 W. Capital Way", "Tacoma", "WA", "USA"));
            employees.Add(new Employee("Roland", "Mendel", "Sales Representative", "722 Moss Bay Blvd.", "Kirkland", "WA", "USA"));
            employees.Add(new Employee("Margaret", "Peacock", "Sales Representative", "4110 Old Redmond Rd.", "Redmond", "WA", "USA"));
            employees.Add(new Employee("Steven", "Buchanan", "Sales Manager", "14 Garrett Hill", "London", "Kirkland", "UK"));
            return employees;
        }
        /// <summary>
        /// Represents the method that handles MergeImageField event.
        /// </summary>
        private static void MergeField_LogoImage(object sender, MergeImageFieldEventArgs args)
        {
            //Binds image from file system during mail merge.
            if (args.FieldName == "Logo")
            {
                string photoFileName = args.FieldValue.ToString();
                //Gets the image from file system.
                FileStream imageStream = new FileStream(Path.GetFullPath(@"Data/" + photoFileName), FileMode.Open, FileAccess.Read);
                args.ImageStream = imageStream;
            }
        }
        private static DataTable GetDataTable()
        {
            //Creates new DataTable instance 
            DataTable table = new DataTable();
            //Add columns in DataTable
            table.Columns.Add("Logo");

            //Add record in new DataRow
            DataRow row = table.NewRow();
            row["Logo"] = "Picture.png";
            table.Rows.Add(row);

            return table;
        }
    }

    /// <summary>
    /// Represents a class to maintain employee details.
    /// </summary>
    public class Employee
    {
        public string FirstName { get; set; }
        public string LastName { get; set; }
        public string Address { get; set; }
        public string City { get; set; }
        public string Region { get; set; }
        public string Country { get; set; }
        public string Title { get; set; }
        public Employee(string firstName, string lastName, string title, string address, string city, string region, string country)
        {
            FirstName = firstName;
            LastName = lastName;
            Title = title;
            Address = address;
            City = city;
            Region = region;
            Country = country;
        }
    }
}
