using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.Collections;
using System.Collections.Generic;
using System.Dynamic;
using System.IO;

namespace Mail_merge_with_dynamic_objects
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
                    //Creates an instance of the MailMergeDataSet.
                    MailMergeDataSet dataSet = new MailMergeDataSet();
                    //Creates the mail merge data table in order to perform mail merge.
                    MailMergeDataTable dataTable = new MailMergeDataTable("Customers", GetCustomers());
                    dataSet.Add(dataTable);
                    dataTable = new MailMergeDataTable("Orders", GetOrders());
                    dataSet.Add(dataTable);
                    List<DictionaryEntry> commands = new List<DictionaryEntry>();
                    //DictionaryEntry contain "Source table" (key) and "Command" (value).
                    DictionaryEntry entry = new DictionaryEntry("Customers", string.Empty);
                    commands.Add(entry);
                    //Retrieves the customer details.
                    entry = new DictionaryEntry("Orders", "CustomerID = %Customers.CustomerID%");
                    commands.Add(entry);
                    //Performs the mail merge operation with the dynamic collection.
                    document.MailMerge.ExecuteNestedGroup(dataSet, commands);
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
        /// Get the customers details to perform mail merge.
        /// </summary>
        private static List<ExpandoObject> GetCustomers()
        {
            List<ExpandoObject> customers = new List<ExpandoObject>();
            customers.Add(GetDynamicCustomer(100, "Robert", "Syncfusion"));
            customers.Add(GetDynamicCustomer(102, "John", "Syncfusion"));
            customers.Add(GetDynamicCustomer(110, "David", "Syncfusion"));
            return customers;
        }
        /// <summary>
        /// Get the order details to perform mail merge
        /// </summary>
		private static List<ExpandoObject> GetOrders()
        {
            List<ExpandoObject> orders = new List<ExpandoObject>();
            orders.Add(GetDynamicOrder(1001, "MSWord", 100));
            orders.Add(GetDynamicOrder(1002, "AdobeReader", 100));
            orders.Add(GetDynamicOrder(1003, "VisualStudio", 102));
            return orders;
        }
        /// <summary>
        /// Generate customer details as dynamic objects.
        /// </summary>
        /// <param name="customerID">Represents an customer id</param>
        /// <param name="customerName">Represents a customer name</param>
        /// <param name="companyName">Represents a company name</param>
		private static dynamic GetDynamicCustomer(int customerID, string customerName, string companyName)
        {
            dynamic dynamicCustomer = new ExpandoObject();
            dynamicCustomer.CustomerID = customerID;
            dynamicCustomer.CustomerName = customerName;
            dynamicCustomer.CompanyName = companyName;
            return dynamicCustomer;
        }
        /// <summary>
        /// Generate order details as dynamic objects.
        /// </summary>
        /// <param name="orderID">Represents an order id</param>
        /// <param name="orderName">Represents an order name</param>
        /// <param name="customerID">Represents customer Id</param>
		private static dynamic GetDynamicOrder(int orderID, string orderName, int customerID)
        {
            dynamic dynamicOrder = new ExpandoObject();
            dynamicOrder.OrderID = orderID;
            dynamicOrder.OrderName = orderName;
            dynamicOrder.CustomerID = customerID;
            return dynamicOrder;
        }
        #endregion
    }
}
