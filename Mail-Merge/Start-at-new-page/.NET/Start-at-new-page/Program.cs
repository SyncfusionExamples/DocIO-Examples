using System.Collections.Generic;
using System.IO;
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;

namespace Start_at_new_page
{
    class Program
    {
        static void Main(string[] args)
        {
            using (FileStream fileStream = new FileStream(Path.GetFullPath(@"Data/Template.docx"), FileMode.Open))
            {
                //Loads an existing Word document.
                using (WordDocument document = new WordDocument(fileStream, FormatType.Automatic))
                {
                    //Gets the invoice details as “IEnumerable” collection.
                    List<Invoice> invoice = GetInvoice();
                    //Creates an instance of “MailMergeDataTable” by specifying mail merge group name and “IEnumerable” collection.
                    MailMergeDataTable dataTable = new MailMergeDataTable("Invoice", invoice);
                    //Enables the flag to start each record in new page.
                    document.MailMerge.StartAtNewPage = true;
                    //Performs Mail merge.
                    document.MailMerge.ExecuteNestedGroup(dataTable);
                    //Creates file stream.
                    using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"Output/Result.docx"), FileMode.Create, FileAccess.ReadWrite))
                    {
                        //Saves the Word document to file stream.
                        document.Save(outputFileStream, FormatType.Docx);
                    }
                }
            }
        }

        #region Helper method
        /// <summary>
        /// Get the data to perform mail merge.
        /// </summary>
        public static List<Invoice> GetInvoice()
        {
            //Creates invoice details.
            List<Invoice> invoices = new List<Invoice>();

            List<Orders> orders = new List<Orders>();
            orders.Add(new Orders("10248", "Vins et alcools Chevalier", "59 rue de l'Abbaye", "Reims", "51100", "France", "VINET", "59 rue de l'Abbaye", "51100", "Reims", "France", "Steven Buchanan", "Vins et alcools Chevalier", "1996-07-04T00:00:00-04:00", "1996-08-01T00:00:00-04:00", "1996-07-16T00:00:00-04:00", "Federal Shipping"));

            List<Order> order = new List<Order>();
            order.Add(new Order("1", "Chai", "14.4", "45", "0.2", "518.4"));
            order.Add(new Order("2", "Boston Crab Meat", "14.7", "40", "0.2", "470.4"));

            List<OrderTotals> orderTotals = new List<OrderTotals>();
            orderTotals.Add(new OrderTotals("440", "32.8", "472.38"));

            invoices.Add(new Invoice(orders, order, orderTotals));

            orders = new List<Orders>();
            orders.Add(new Orders("10249", "Toms Spezialitäten", "Luisenstr. 48", "Münster", "51100", "Germany", "TOMSP", "Luisenstr. 48", "51100", "Münster", "Germany", "Michael Suyama", "Toms Spezialitäten", "1996-07-04T00:00:00-04:00", "1996-08-01T00:00:00-04:00", "1996-07-16T00:00:00-04:00", "Speedy Express"));

            order = new List<Order>();
            order.Add(new Order("1", "Chai", "18", "45", "0.2", "618.4"));
            order.Add(new Order("4", "Alice Mutton", "39", "100", "0", "3900"));

            orderTotals = new List<OrderTotals>();
            orderTotals.Add(new OrderTotals("1863.4", "11.61", "1875.01"));

            invoices.Add(new Invoice(orders, order, orderTotals));

            orders = new List<Orders>();
            orders.Add(new Orders("10250", "Hanari Carnes", "Rua do Paço, 67", "Rio de Janeiro", "05454-876", "Brazil", "VINET", "Rua do Paço, 67", "51100", "Rio de Janeiro", "Brazil", "Margaret Peacock", "Hanari Carnes", "1996-07-04T00:00:00-04:00", "1996-08-01T00:00:00-04:00", "1996-07-16T00:00:00-04:00", "United Package"));

            order = new List<Order>();
            order.Add(new Order("65", "Louisiana Fiery Hot Pepper Sauce", "16.8", "15", "0.15", "214.2"));
            order.Add(new Order("51", "Manjimup Dried Apples", "42.4", "35", "0.15", "1261.4"));

            orderTotals = new List<OrderTotals>();
            orderTotals.Add(new OrderTotals("1552.6", "65.83", "1618.43"));

            invoices.Add(new Invoice(orders, order, orderTotals));

            return invoices;
        }
        #endregion
    }

    #region Helper classes
    /// <summary>
    /// Represents a class to maintain invoice details.
    /// </summary>
    public class Invoice
    {
        #region Fields
        private List<Orders> m_orders;
        private List<Order> m_order;
        private List<OrderTotals> m_orderTotal;
        #endregion

        #region Properties
        public List<Orders> Orders
        {
            get { return m_orders; }
            set { m_orders = value; }
        }
        public List<Order> Order
        {
            get { return m_order; }
            set { m_order = value; }
        }
        public List<OrderTotals> OrderTotals
        {
            get { return m_orderTotal; }
            set { m_orderTotal = value; }
        }
        #endregion

        #region Constructor
        public Invoice(List<Orders> orders, List<Order> order, List<OrderTotals> orderTotals)
        {
            Orders = orders;
            Order = order;
            OrderTotals = orderTotals;
        }
        #endregion
    }
    /// <summary>
    /// Represents a class to maintain orders details.
    /// </summary>
    public class Orders
    {
        #region Fields
        private string m_orderID;
        private string m_shipName;
        private string m_shipAddress;
        private string m_shipCity;
        private string m_shipPostalCode;
        private string m_shipCountry;
        private string m_customerID;
        private string m_address;
        private string m_postalCode;
        private string m_city;
        private string m_country;
        private string m_salesPerson;
        private string m_customersCompanyName;
        private string m_orderDate;
        private string m_requiredDate;
        private string m_shippedDate;
        private string m_shippersCompanyName;
        #endregion

        #region Properties
        public string ShipName
        {
            get { return m_shipName; }
            set { m_shipName = value; }
        }
        public string ShipAddress
        {
            get { return m_shipAddress; }
            set { m_shipAddress = value; }
        }
        public string ShipCity
        {
            get { return m_shipCity; }
            set { m_shipCity = value; }
        }
        public string ShipPostalCode
        {
            get { return m_shipPostalCode; }
            set { m_shipPostalCode = value; }
        }
        public string PostalCode
        {
            get { return m_postalCode; }
            set { m_postalCode = value; }
        }
        public string ShipCountry
        {
            get { return m_shipCountry; }
            set { m_shipCountry = value; }
        }
        public string CustomerID
        {
            get { return m_customerID; }
            set { m_customerID = value; }
        }
        public string Customers_CompanyName
        {
            get { return m_customersCompanyName; }
            set { m_customersCompanyName = value; }
        }
        public string Address
        {
            get { return m_address; }
            set { m_address = value; }
        }
        public string City
        {
            get { return m_city; }
            set { m_city = value; }
        }
        public string Country
        {
            get { return m_country; }
            set { m_country = value; }
        }
        public string Salesperson
        {
            get { return m_salesPerson; }
            set { m_salesPerson = value; }
        }
        public string OrderID
        {
            get { return m_orderID; }
            set { m_orderID = value; }
        }
        public string OrderDate
        {
            get { return m_orderDate; }
            set { m_orderDate = value; }
        }
        public string RequiredDate
        {
            get { return m_requiredDate; }
            set { m_requiredDate = value; }
        }
        public string ShippedDate
        {
            get { return m_shippedDate; }
            set { m_shippedDate = value; }
        }
        public string Shippers_CompanyName
        {
            get { return m_shippersCompanyName; }
            set { m_shippersCompanyName = value; }
        }
        #endregion

        #region Constructor
        public Orders(string orderID, string shipName, string shipAddress, string shipCity,
         string shipPostalCode, string shipCountry, string customerID, string address,
         string postalCode, string city, string country, string salesPerson, string customersCompanyName,
         string orderDate, string requiredDate, string shippedDate, string shippersCompanyName)
        {
            OrderID = orderID;
            ShipName = shipName;
            ShipAddress = shipAddress;
            ShipCity = shipCity;
            ShipPostalCode = shipPostalCode;
            ShipCountry = shipCountry;
            CustomerID = customerID;
            Address = address;
            PostalCode = postalCode;
            City = city;
            Country = country;
            Salesperson = salesPerson;
            Customers_CompanyName = customersCompanyName;
            OrderDate = orderDate;
            RequiredDate = requiredDate;
            ShippedDate = shippedDate;
            Shippers_CompanyName = shippersCompanyName;
        }
        #endregion
    }
    /// <summary>
    /// Represents a class to maintain order details.
    /// </summary>
    public class Order
    {
        #region Fields
        private string m_productID;
        private string m_productName;
        private string m_unitPrice;
        private string m_quantity;
        private string m_discount;
        private string m_extendedPrice;
        #endregion

        #region Properties
        public string ProductID
        {
            get { return m_productID; }
            set { m_productID = value; }
        }
        public string ProductName
        {
            get { return m_productName; }
            set { m_productName = value; }
        }
        public string UnitPrice
        {
            get { return m_unitPrice; }
            set { m_unitPrice = value; }
        }
        public string Quantity
        {
            get { return m_quantity; }
            set { m_quantity = value; }
        }
        public string Discount
        {
            get { return m_discount; }
            set { m_discount = value; }
        }
        public string ExtendedPrice
        {
            get { return m_extendedPrice; }
            set { m_extendedPrice = value; }
        }
        #endregion

        #region Constructor       
        public Order(string productID, string productName, string unitPrice, string quantity,
         string discount, string extendedPrice)
        {
            ProductID = productID;
            ProductName = productName;
            UnitPrice = unitPrice;
            Quantity = quantity;
            Discount = discount;
            ExtendedPrice = extendedPrice;
        }
        #endregion
    }
    /// <summary>
    /// Represents a class to maintain order totals details.
    /// </summary>
    public class OrderTotals
    {
        #region Fields
        private string m_subTotal;
        private string m_freight;
        private string m_total;
        #endregion

        #region Properties
        public string Subtotal
        {
            get { return m_subTotal; }
            set { m_subTotal = value; }
        }
        public string Freight
        {
            get { return m_freight; }
            set { m_freight = value; }
        }
        public string Total
        {
            get { return m_total; }
            set { m_total = value; }
        }
        #endregion

        #region Constructor       
        public OrderTotals(string subTotal, string freight, string total)
        {
            Subtotal = subTotal;
            Freight = freight;
            Total = total;
        }
        #endregion
    }
    #endregion
}
