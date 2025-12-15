using Syncfusion.DocIO.DLS;
using System.Dynamic;
using System.Xml;

namespace Generate_invoices_from_xml_data
{
    class Program
    {
        public static void Main(string[] args)
        {
            // Load the template Word document
            WordDocument document = new WordDocument(Path.GetFullPath(@"Data/Template.docx"));
            // start each record at new page
            document.MailMerge.StartAtNewPage = true;
            // Perform mail merge using relational XML data
            document.MailMerge.ExecuteNestedGroup(GetRelationalData());
            // Save the result Word document.
            document.Save(Path.GetFullPath("../../../Output/Output.docx"));
            // Close the Word document
            document.Close();
        }
        #region Helper Method
        /// <summary>
        /// Retrieves relational invoice data from XML and converts it into a MailMergeDataTable.
        /// </summary>
        private static MailMergeDataTable GetRelationalData()
        {            
            // Load the XML data file
            Stream xmlStream = File.OpenRead(Path.GetFullPath(@"Data/InvoiceDetails.xml"));
            XmlDocument xmlDocument = new XmlDocument();
            xmlDocument.Load(xmlStream);
            xmlStream.Dispose();

            ExpandoObject customerDetails = new ExpandoObject();
            GetDataAsExpandoObject((xmlDocument as XmlNode).LastChild, ref customerDetails);
            // Convert dynamic object into dictionary
            IDictionary<string, object> customerDict = customerDetails as IDictionary<string, object>;
            // Get the "Invoices" list
            List<ExpandoObject> invoicesList = customerDict["Invoices"] as List<ExpandoObject>;
            // Take the first item in that list
            IDictionary<string, object> firstInvoiceGroup = invoicesList[0] as IDictionary<string, object>;
            // Get the "Invoice" list from that group
            List<ExpandoObject>  invoices = firstInvoiceGroup["Invoice"] as List<ExpandoObject>;
            // Create MailMergeDataTable with group name "Invoices"
            MailMergeDataTable dataTable = new MailMergeDataTable("Invoices", invoices);
            return dataTable;
        }
        /// <summary>
        /// Gets the data as ExpandoObject.
        /// </summary>
        /// <param name="node">The current XML node being processed.</param>
        /// <param name="dynamicObject">The dynamic object to populate with node data.</param>
        /// <returns></returns>
        private static void GetDataAsExpandoObject(XmlNode node, ref ExpandoObject dynamicObject)
        {
            if (node.InnerText == node.InnerXml)
                // Leaf node: store text value
                dynamicObject.TryAdd(node.LocalName, node.InnerText);
            else
            {
                // Handle child nodes
                List<ExpandoObject> childObjects;
                // If the tag already exists, reuse the list; otherwise create a new one
                if ((dynamicObject as IDictionary<string, object>).ContainsKey(node.LocalName))
                    childObjects = (dynamicObject as IDictionary<string, object>)[node.LocalName] as List<ExpandoObject>;
                else
                {
                    childObjects = new List<ExpandoObject>();
                    dynamicObject.TryAdd(node.LocalName, childObjects);
                }
                // Create a new child object for the current node
                ExpandoObject childObject = new ExpandoObject();
                // Recursively process all child nodes
                foreach (XmlNode childNode in (node as XmlNode).ChildNodes)
                {
                    GetDataAsExpandoObject(childNode, ref childObject);
                }
                // Add the processed child object to the list
                childObjects.Add(childObject);
            }
        }
        #endregion
    }
}

