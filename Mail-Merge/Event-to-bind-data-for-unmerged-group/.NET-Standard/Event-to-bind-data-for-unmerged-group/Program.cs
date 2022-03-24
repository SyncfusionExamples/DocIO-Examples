using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System;
using System.Collections.Generic;
using System.IO;

namespace Event_to_bind_data_for_unmerged_group
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
					//Sets “ClearFields” to true to remove empty mail merge fields from document.
					document.MailMerge.ClearFields = false;
					//Uses the mail merge event to clear the unmerged group field while perform mail merge execution.
					document.MailMerge.BeforeClearGroupField += new BeforeClearGroupFieldEventHandler(BeforeClearFields);
					//Gets the employee details as “IEnumerable” collection.
					List<Employees> employeeList = GetEmployees();
					//Creates an instance of “MailMergeDataTable” by specifying mail merge group name and “IEnumerable” collection.
					MailMergeDataTable dataTable = new MailMergeDataTable("Employees", employeeList);
					//Performs Mail merge
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

        #region Helper Methods
        /// <summary>
        /// Represents the method that handles BeforeClearGroupField event.
        /// </summary>
        private static void BeforeClearFields(object sender, BeforeClearGroupFieldEventArgs args)
        {
            if (!args.HasMappedGroupInDataSource)
            {
                //Gets the Current unmerged group name from the event argument.
                string[] groupName = args.GroupName.Split(':');
                if (groupName[groupName.Length - 1] == "Orders")
                {
                    string[] fields = args.FieldNames;
                    List<OrderDetails> orderList = GetOrders();
                    //Binds the data to the unmerged fields in group as alternative values.
                    args.AlternateValues = orderList;
                }
                else
                    //If group value is empty, you can set whether the unmerged merge group field can be clear or not.
                    args.ClearGroup = true;
            }
        }
        /// <summary>
        /// Get the order details to perform mail merge
        /// </summary>
        private static List<OrderDetails> GetOrders()
        {
            List<OrderDetails> orders = new List<OrderDetails>();
            orders.Add(new OrderDetails("10952", new DateTime(1998, 3, 16), new DateTime(1998, 3, 24), new DateTime(1998, 2, 12)));
            return orders;
        }
        /// <summary>
        /// Get the employee details to perform mail merge.
        /// </summary>
        /// <returns></returns>
        public static List<Employees> GetEmployees()
        {
            List<OrderDetails> orders = new List<OrderDetails>();
            orders.Add(new OrderDetails("10835", new DateTime(1998, 1, 15), new DateTime(1998, 1, 21), new DateTime(1998, 2, 12)));
            List<CustomerDetails> customerDetails = new List<CustomerDetails>();
            customerDetails.Add(new CustomerDetails("Maria Anders", "Alfreds Futterkiste", "Berlin", "Germany", orders));
            customerDetails.Add(new CustomerDetails("Thomas Hardy", "Around the Horn", "London", "UK", null));
            List<Employees> employees = new List<Employees>();
            employees.Add(new Employees("Nancy", "Smith", "1001", "505 - 20th Ave. E. Apt. 2A", "Seattle WA", "USA", customerDetails));
            return employees;
        }
        #endregion
    }

    #region Helper Class
    /// <summary>
    /// Represents a class to maintain employee details.
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
    /// Represents a class to maintain customer details.
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
    /// Represents a class to maintain order details.
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
