using Syncfusion.DocIO.DLS;
using Syncfusion.DocIO;
using System.Dynamic;
using System.Xml;

namespace Sort_expandoobject_collection_before_mailmerge
{
    class Program
    {
        static void Main(string[] args)
        {
            using (FileStream fileStream = new FileStream(Path.GetFullPath(@"../../../Data/Template.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                using (WordDocument document = new WordDocument(fileStream, FormatType.Docx))
                {
                    var dataTable = new MailMergeDataTable("Customers", GetCustomers(Path.GetFullPath(@"../../../Data/Data.xml")));
                    document.MailMerge.ExecuteNestedGroup(dataTable);

                    using (FileStream outStream = new FileStream(Path.GetFullPath(@"../../../Output/output.docx"), FileMode.Create, FileAccess.Write))
                    {
                        document.Save(outStream, FormatType.Docx);
                    }
                }
            }
        }

        /// <summary>
        /// Reads customer information from the XML file, creates a collection of dynamic customer objects,
        /// sorts all hierarchy levels, and returns the prepared customer list.
        /// </summary>
        private static List<ExpandoObject> GetCustomers(string xmlPath)
        {
            List<ExpandoObject> customers = new List<ExpandoObject>();
            using (XmlReader reader = XmlReader.Create(xmlPath))
            {
                while (reader.Read())
                {
                    if (reader.IsStartElement() && reader.Name == "Customer")
                    {
                        string customerId = null;
                        string contactName = null;
                        string companyName = null;
                        List<ExpandoObject> orders = new List<ExpandoObject>();
                        while (reader.Read())
                        {
                            if (reader.IsStartElement())
                            {
                                switch (reader.Name)
                                {
                                    case "CustomerID":
                                        customerId = reader.ReadElementContentAsString();
                                        break;
                                    case "ContactName":
                                        contactName = reader.ReadElementContentAsString();
                                        break;
                                    case "CompanyName":
                                        companyName = reader.ReadElementContentAsString();
                                        break;
                                    case "Orders":
                                        orders = GetOrders(reader);
                                        break;
                                }
                            }
                            else if (reader.NodeType == XmlNodeType.EndElement && reader.Name == "Customer")
                            {
                                customers.Add(GetDynamicCustomer(
                                    int.Parse(customerId),
                                    contactName,
                                    companyName,
                                    orders
                                ));
                                break;
                            }
                        }
                    }
                }
            }
            // Sort all levels of expandOobject.
            SortAllLevels(customers);
            return customers;
        }

        /// <summary>
        /// Reads order and order item details from the XML reader and creates a collection
        /// of dynamic order objects with their corresponding nested order items.
        /// </summary>
        private static List<ExpandoObject> GetOrders(XmlReader reader)
        {
            List<ExpandoObject> orders = new List<ExpandoObject>();
            while (reader.Read())
            {
                if (reader.NodeType == XmlNodeType.Element && reader.Name == "Order")
                {
                    dynamic order = new ExpandoObject();
                    var orderDict = (IDictionary<string, object>)order;
                    List<ExpandoObject> orderItems = new List<ExpandoObject>();

                    while (reader.Read())
                    {
                        if (reader.IsStartElement())
                        {
                            switch (reader.Name)
                            {
                                case "OrderID":
                                    order.OrderID = reader.ReadElementContentAsInt();
                                    break;
                                case "OrderName":
                                    order.OrderName = reader.ReadElementContentAsString();
                                    break;
                                case "CustomerID":
                                    order.CustomerID = reader.ReadElementContentAsInt();
                                    break;
                                case "OrderItems":
                                    while (reader.Read())
                                    {
                                        if (reader.IsStartElement() && reader.Name == "Item")
                                        {
                                            dynamic item = new ExpandoObject();
                                            var itemDict = (IDictionary<string, object>)item;
                                            while (reader.Read())
                                            {
                                                if (reader.IsStartElement())
                                                {
                                                    string element = reader.Name;
                                                    string value = reader.ReadElementContentAsString();
                                                    itemDict[element] = int.TryParse(value, out int intVal) ? intVal : (object)value;
                                                }
                                                else if (reader.NodeType == XmlNodeType.EndElement && reader.Name == "Item")
                                                {
                                                    orderItems.Add(item);
                                                    break;
                                                }
                                            }
                                        }
                                        else if (reader.NodeType == XmlNodeType.EndElement && reader.Name == "OrderItems")
                                        {
                                            break;
                                        }
                                    }
                                    order.OrderItems = orderItems;
                                    break;
                            }
                        }
                        else if (reader.NodeType == XmlNodeType.EndElement && reader.Name == "Order")
                        {
                            break;
                        }
                    }
                    orders.Add(order);
                }
                else if (reader.NodeType == XmlNodeType.EndElement && reader.Name == "Orders")
                {
                    break;
                }
            }
            return orders;
        }

        /// <summary>
        /// Creates and returns a dynamic customer object with customer details
        /// and its associated orders collection.
        /// </summary>
        private static dynamic GetDynamicCustomer(int customerID, string customerName, string companyName, List<ExpandoObject> orders)
        {
            dynamic dynamicCustomer = new ExpandoObject();
            dynamicCustomer.CustomerID = customerID;
            dynamicCustomer.CustomerName = customerName;
            dynamicCustomer.CompanyName = companyName;
            dynamicCustomer.Orders = orders;
            return dynamicCustomer;
        }

        /// <summary>
        /// Initiates recursive sorting for the provided ExpandoObject collection
        /// across all hierarchy levels.
        /// </summary>
        public static void SortAllLevels(List<ExpandoObject> dynamicList)
        {
            SortRecursively(dynamicList);
        }

        /// <summary>
        /// Recursively sorts the current ExpandoObject collection based on the first
        /// comparable property and processes any nested ExpandoObject collections.
        /// </summary>
        private static void SortRecursively(List<ExpandoObject> currentLevelList)
        {
            if (currentLevelList == null || currentLevelList.Count == 0) return;

            var firstItemDictionary = (IDictionary<string, object>)currentLevelList[0];

            string keyToSortBy = firstItemDictionary
                .FirstOrDefault(entry => entry.Value is string || entry.Value is int || entry.Value is IComparable)
                .Key;

            if (!string.IsNullOrEmpty(keyToSortBy))
            {
                currentLevelList.Sort((firstItem, secondItem) =>
                {
                    var firstItemDict = (IDictionary<string, object>)firstItem;
                    var secondItemDict = (IDictionary<string, object>)secondItem;

                    object firstValue = firstItemDict.ContainsKey(keyToSortBy) ? firstItemDict[keyToSortBy] : null;
                    object secondValue = secondItemDict.ContainsKey(keyToSortBy) ? secondItemDict[keyToSortBy] : null;

                    return Comparer<object>.Default.Compare(firstValue, secondValue);
                });
            }

            // Recursively sort any nested lists within the current level
            foreach (ExpandoObject currentItem in currentLevelList)
            {
                var currentItemDict = (IDictionary<string, object>)currentItem;

                foreach (var property in currentItemDict)
                {
                    if (property.Value is List<ExpandoObject> nestedList)
                    {
                        SortRecursively(nestedList);
                    }
                }
            }
        }
    }
}