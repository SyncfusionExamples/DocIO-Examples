using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System;
using System.Collections.Generic;
using System.IO;

namespace Remove_unmerged_rows_during_mail_merge
{
    class Program
    {
        static void Main(string[] args)
        {
            using (FileStream fileStream = new FileStream(Path.GetFullPath(@"Data/Template.docx"), FileMode.Open, FileAccess.ReadWrite))
            {
                //Loads an existing Word document into DocIO instance.
                using (WordDocument document = new WordDocument(fileStream, FormatType.Automatic))
                {
                    //Gets the employee details as IEnumerable collection.
                    List<Employee> employeeList = GetEmployees();

                    //Event to do manipulations when unmerged field occurs in the Word document
                    document.MailMerge.BeforeClearField += RemoveRowOfDataNotInDataSource;

                    //Event to remove row when fields have empty value but defined in Datasource  
                    document.MailMerge.MergeField += RemoveRowsOfEmptyValue;

                    //Creates an instance of MailMergeDataTable by specifying MailMerge group name and IEnumerable collection.
                    MailMergeDataTable dataSource = new MailMergeDataTable("Employee", employeeList);
                    //Performs Mail merge.
                    document.MailMerge.ExecuteGroup(dataSource);
                    //Creates file stream.
                    using (FileStream outputStream = new FileStream(Path.GetFullPath(@"Output/Result.docx"), FileMode.Create, FileAccess.ReadWrite))
                    {
                        //Saves the Word document to file stream.
                        document.Save(outputStream, FormatType.Docx);
                    }
                }
            }
        }

        /// <summary>
        /// Remove a row if data is not defined in Data source
        /// </summary>
        private static void RemoveRowOfDataNotInDataSource(object sender, BeforeClearFieldEventArgs args)
        {
            if (args.GroupName == "Employee" && !args.HasMappedFieldInDataSource)
            {
                WParagraph paragraph = args.CurrentMergeField.OwnerParagraph;
                if (paragraph.IsInCell)
                {
                    WTableRow tableRow = paragraph.Owner.Owner as WTableRow;
                    (tableRow.Owner as WTable).Rows.Remove(tableRow);
                }
            }
        }

        /// <summary>
        /// Event handler to remove the row of empty or null value fields
        /// </summary>
        private static void RemoveRowsOfEmptyValue(object sender, MergeFieldEventArgs args)
        {
            if (args.GroupName == "Employee" && (args.FieldValue == DBNull.Value ||  args.FieldValue.ToString() == string.Empty))
            {
                WParagraph paragraph = args.CurrentMergeField.OwnerParagraph;
                if (paragraph.IsInCell)
                {
                    WTableRow tableRow = paragraph.Owner.Owner as WTableRow;
                    (tableRow.Owner as WTable).Rows.Remove(tableRow);
                }
            }
        }

        /// <summary>
        /// Gets the employee details to perform mail merge.
        /// </summary>
        public static List<Employee> GetEmployees()
        {
            List<Employee> employees = new List<Employee>
            {
                new Employee("Nancy", "", "722 Moss Bay Blvd.", "USA"),
                new Employee("Andrew", "12/12/1988", "", "USA"),
                new Employee("Roland", "03/22/1992", "722 Moss Bay Blvd.", "USA"),
                new Employee("Margaret", "07/19/1980", "", "USA"),
                new Employee("Steven", "09/30/1995", "14 Garrett Hill", "")
            };
            return employees;
        }
    }

    /// <summary>
    /// Represents a class to maintain employee details.
    /// </summary>
    public class Employee
    {
        public string Name { get; set; }
        public string DOB { get; set; }  
        public string Address { get; set; }
        public string Country { get; set; }

        public Employee(string name, string dob, string address, string country)
        {
            Name = name;
            DOB = dob;
            Address = address;
            Country = country;
        }
    }
}
