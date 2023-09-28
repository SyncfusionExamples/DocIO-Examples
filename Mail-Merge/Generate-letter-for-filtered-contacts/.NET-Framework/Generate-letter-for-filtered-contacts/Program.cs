using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.Data;
using System.IO;

namespace Group_customers_based_on_products
{
    class Program
    {
        static void Main(string[] args)
        {
            //Creates new Word document instance for Word processing.
            using (WordDocument document = new WordDocument())
            {
                //Opens the Word template document.
                document.Open(Path.GetFullPath(@"../../LetterTemplate.docx"), FormatType.Docx);
                //Get the contact details.
                DataTable table = GetContacts();
                //Creates a DataView for DataTable.
                DataView dataView = new DataView(table);
                //Filter the customers from USA.
                dataView.RowFilter = "Country = 'USA'";
                //Perform mail merge.
                document.MailMerge.ExecuteGroup(dataView);
                //Saves the Word document.
                document.Save(Path.GetFullPath(@"../../Sample.docx"), FormatType.Docx);
            }
        }
        #region Helper methods
        /// <summary>
        /// Gets the data to perform mail merge.
        /// </summary>
        /// <returns></returns>
        private static DataTable GetContacts()
        {
            //Creates new DataTable instance. 
            DataTable table = new DataTable("Contacts");
            //Add columns for the DataTable.
            table.Columns.Add("ContactName");
            table.Columns.Add("CompanyName");
            table.Columns.Add("Address");
            table.Columns.Add("City");
            table.Columns.Add("Country");
            table.Columns.Add("Phone");

            //Add records in DataTable.
            DataRow row = table.NewRow();
            row["ContactName"] = "Fran Wilson";
            row["CompanyName"]= "Lonesome Pine Restaurant";
            row["Address"] = "89 Chiaroscuro Rd.";
            row["City"]= "Portland OR";
            row["Country"] = "USA";
            row["Phone"] = "(503) 555-9573";
            table.Rows.Add(row);

            row = table.NewRow();
            row["ContactName"] = "John Steel";
            row["CompanyName"] = "Lazy K Kountry Store";
            row["Address"] = "12 Orchestra Terrace";
            row["City"] = "Walla Walla WA";
            row["Country"] = "USA";
            row["Phone"] = "(509) 555-7969";
            table.Rows.Add(row);

            row = table.NewRow();
            row["ContactName"] = "Helvetius Nagy";
            row["CompanyName"] = "Trail's Head Gourmet Provisioners";
            row["Address"] = "722 DaVinci Blvd.";
            row["City"] = "Kirkland WA";
            row["Country"] = "USA";
            row["Phone"] = "(206) 555-8257";
            table.Rows.Add(row);

            row = table.NewRow();
            row["ContactName"] = "Victoria Ashworth";
            row["CompanyName"] = "B's Beverages";
            row["Address"] = "Fauntleroy Circus";
            row["City"] = "London";
            row["Country"] = "UK";
            row["Phone"] = "(171) 555-1212";
            table.Rows.Add(row);

            row = table.NewRow();
            row["ContactName"] = "Hanna Moos";
            row["CompanyName"] = "Blauer See Delikatessen";
            row["Address"] = "Forsterstr. 57";
            row["City"] = "Mannheim";
            row["Country"] = "Germany";
            row["Phone"] = "0621-08460";
            table.Rows.Add(row);
           
            row = table.NewRow();
            row["ContactName"] = "Howard Snyder";
            row["CompanyName"] = "Great Lakes Food Market";
            row["Address"] = "2732 Baker Blvd.";
            row["City"] = "Eugene";
            row["Country"] = "USA";
            row["Phone"] = "(503) 555-7555";
            table.Rows.Add(row);

            return table;
        }
        #endregion
    }
}
