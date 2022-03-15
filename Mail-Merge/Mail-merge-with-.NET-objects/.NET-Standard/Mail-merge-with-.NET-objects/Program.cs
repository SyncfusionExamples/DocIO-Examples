using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System;
using System.Collections.Generic;
using System.IO;

namespace Mail_merge_with_.NET_objects
{
    class Program
    {
        static void Main(string[] args)
        {
            using (FileStream fileStream = new FileStream(Path.GetFullPath(@"../../../Data/EmployeesTemplate.docx"), FileMode.Open, FileAccess.ReadWrite))
            {
                //Loads an existing Word document into DocIO instance.
                using (WordDocument document = new WordDocument(fileStream, FormatType.Automatic))
                {
                    //Gets the employee details as IEnumerable collection.
                    List<Employee> employeeList = GetEmployees();
                    //Creates an instance of MailMergeDataTable by specifying MailMerge group name and IEnumerable collection.
                    MailMergeDataTable dataSource = new MailMergeDataTable("Employees", employeeList);
                    //Uses the mail merge events handler for image fields.
                    document.MailMerge.MergeImageField += new MergeImageFieldEventHandler(MergeField_EmployeeImage);
                    //Performs Mail merge.
                    document.MailMerge.ExecuteGroup(dataSource);
                    //Creates file stream.
                    using (FileStream outputStream = new FileStream(Path.GetFullPath(@"../../../Result.docx"), FileMode.Create, FileAccess.ReadWrite))
                    {
                        //Saves the Word document to file stream.
                        document.Save(outputStream, FormatType.Docx);
                    }
                }
            }
        }

        public static List<Employee> GetEmployees()
        {
            List<Employee> employees = new List<Employee>();
            employees.Add(new Employee("Nancy", "Smith", "Sales Representative", "505 - 20th Ave. E. Apt. 2A,", "Seattle", "WA", "USA", "Nancy.png"));
            employees.Add(new Employee("Andrew", "Fuller", "Vice President, Sales", "908 W. Capital Way", "Tacoma", "WA", "USA", "Andrew.png"));
            employees.Add(new Employee("Roland", "Mendel", "Sales Representative", "722 Moss Bay Blvd.", "Kirkland", "WA", "USA", "Janet.png"));
            employees.Add(new Employee("Margaret", "Peacock", "Sales Representative", "4110 Old Redmond Rd.", "Redmond", "WA", "USA", "Margaret.png"));
            employees.Add(new Employee("Steven", "Buchanan", "Sales Manager", "14 Garrett Hill", "London", string.Empty, "UK", "Steven.png"));
            return employees;
        }

        private static void MergeField_EmployeeImage(object sender, MergeImageFieldEventArgs args)
        {
            //Binds image from file system during mail merge.
            if (args.FieldName == "Photo")
            {
                string photoFileName = args.FieldValue.ToString();
                //Gets the image from file system.
                FileStream imageStream = new FileStream(Path.GetFullPath(@"../../../Data/" + photoFileName), FileMode.Open, FileAccess.Read);
                args.ImageStream = imageStream;
            }
        }
    }

    public class Employee
    {
        public string FirstName { get; set; }
        public string LastName { get; set; }
        public string Address { get; set; }
        public string City { get; set; }
        public string Region { get; set; }
        public string Country { get; set; }
        public string Title { get; set; }
        public string Photo { get; set; }
        public Employee(string firstName, string lastName, string title, string address, string city, string region, string country, string photoFilePath)
        {
            FirstName = firstName;
            LastName = lastName;
            Title = title;
            Address = address;
            City = city;
            Region = region;
            Country = country;
            Photo = photoFilePath;
        }
    }
}
