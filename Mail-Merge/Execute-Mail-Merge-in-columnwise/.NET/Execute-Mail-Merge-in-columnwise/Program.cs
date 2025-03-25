using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.Collections.Generic;
using System.IO;

namespace Execute_Mail_Merge_in_columnwise
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
                    //Gets the customer details as “IEnumerable” collection
                    List<CustomerDetail> customerDetails = GetCustomerDetails();
                    //Creates an instance of “MailMergeDataTable” by specifying mail merge group name and “IEnumerable” collection.
                    MailMergeDataTable dataTable = new MailMergeDataTable("CustomerDetails", customerDetails);
                    //Removes the empty paragraph, if the field not have value.
                    document.MailMerge.RemoveEmptyParagraphs = true;
                    //Performs Mail merge
                    document.MailMerge.ExecuteGroup(dataTable);

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
        /// Create a data to execute mail merge in Word document.
        /// </summary>
        /// <returns></returns>
        public static List<CustomerDetail> GetCustomerDetails()
        {
            List<CustomerDetail> customers = new List<CustomerDetail>();
            customers.Add(new CustomerDetail("Sathish Kumar", "F3, Bharath PST Castle,", "14, Renga Nagar, Srirangam", "Trichy,", "Tamilnadu", string.Empty));
            customers.Add(new CustomerDetail("Swathi", "No12, DN Aparments", "15, Thillai ganga nagar", "Chennai,", "Tamilnadu", string.Empty));
            customers.Add(new CustomerDetail("Brent", "#12, London Road", string.Empty, string.Empty, "Oxford,", "United Kingdom"));
            customers.Add(new CustomerDetail("Mani", "#12, Steve Lane,", string.Empty, string.Empty, "Dublin,", "Ireland"));
            return customers;
        }
    }
    /// <summary>
    /// Represents a class to maintain customer details.
    /// </summary>
    public class CustomerDetail
    {
        public string Name { get; set; }
        public string Address1 { get; set; }
        public string Address2 { get; set; }
        public string City { get; set; }
        public string Region { get; set; }
        public string Country { get; set; }

        public CustomerDetail(string name, string address1, string address2, string city, string region, string country)
        {
            Name = name;
            Address1 = address1;
            Address2 = address2;
            City = city;
            Region = region;
            Country = country;
        }
    }
}
