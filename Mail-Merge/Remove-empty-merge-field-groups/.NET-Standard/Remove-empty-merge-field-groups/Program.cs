using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System;
using System.Collections.Generic;
using System.IO;

namespace Remove_empty_merge_field_groups
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
                    //Gets the employee details as “IEnumerable” collection.
                    List<Employees> employeeList = GetEmployees();
                    //Creates an instance of “MailMergeDataTable” by specifying mail merge group name and “IEnumerable” collection.
                    MailMergeDataTable dataTable = new MailMergeDataTable("Employees", employeeList);
                    //Enable the flag to remove empty groups which contain empty merge fields.
                    document.MailMerge.RemoveEmptyGroup = true;
                    //Performs Mail merge.
                    document.MailMerge.ExecuteNestedGroup(dataTable);
                    //Creates file stream.
                    using (FileStream outputStream = new FileStream(Path.GetFullPath(@"../../../Result.docx"), FileMode.Create, FileAccess.ReadWrite))
                    {
                        //Saves the Word document to file stream.
                        document.Save(outputStream, FormatType.Docx);
                    }
                }
            }
        }

        #region Helper methods
        /// <summary>
        /// Represents the method to get employees details.
        /// </summary>
        public static List<Employees> GetEmployees()
        {
            //Create order details.
            List<OrderDetails> orders = new List<OrderDetails>();
            orders.Add(new OrderDetails("10835", new DateTime(2015, 1, 5), new DateTime(2015, 1, 12), new DateTime(2015, 1, 21)));
            //Create customer details.
            List<CustomerDetails> customerDetails = new List<CustomerDetails>();
            customerDetails.Add(new CustomerDetails("Maria Anders", "Maria Anders", "Berlin", "Germany", orders));
            customerDetails.Add(new CustomerDetails("Andy", "Bernard", "Berlin", "Germany", null));
            //Create employee details.
            List<Employees> employees = new List<Employees>();
            employees.Add(new Employees("Nancy", "Smith", "1", "505 - 20th Ave. E. Apt. 2A,", "Seattle", "USA", customerDetails));
            return employees;
        }
        #endregion
    }

    #region Helper Class
    /// <summary>
    /// Represents a class to hold the employee details.
    /// </summary>
    public class Employees
    {
        public string FirstName { get; set; }
        public string LastName { get; set; }
        public string EmployeeID { get; set; }
        public string Address { get; set; }
        public string City { get; set; }
        public string Country { get; set; }
        public List<CustomerDetails> Customers { get; set; }
        public Employees(string firstName, string lastName, string employeeId, string address, string city, string country, List<CustomerDetails> customers)
        {
            FirstName = firstName;
            LastName = lastName;
            Address = address;
            EmployeeID = employeeId;
            City = city;
            Country = country;
            Customers = customers;
        }
    }
    /// <summary>
    /// Represents a class to hold the customer details.
    /// </summary>
    public class CustomerDetails
    {
        public string ContactName { get; set; }
        public string CompanyName { get; set; }
        public string City { get; set; }
        public string Country { get; set; }
        public List<OrderDetails> Orders { get; set; }
        public CustomerDetails(string contactName, string companyName, string city, string country, List<OrderDetails> orders)
        {
            ContactName = contactName;
            CompanyName = companyName;
            City = city;
            Country = country;
            Orders = orders;
        }
    }
    /// <summary>
    /// Represents a class to hold the order details.
    /// </summary>
    public class OrderDetails
    {
        public string OrderID { get; set; }
        public DateTime OrderDate { get; set; }
        public DateTime ShippedDate { get; set; }
        public DateTime RequiredDate { get; set; }
        public OrderDetails(string orderId, DateTime orderDate, DateTime shippedDate, DateTime requiredDate)
        {
            OrderID = orderId;
            OrderDate = orderDate;
            ShippedDate = shippedDate;
            RequiredDate = requiredDate;
        }
    }
    #endregion
}
