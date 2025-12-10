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
            // To start each record at new page, so enable this property
            document.MailMerge.StartAtNewPage = true;
            // Perform the mail merge for the group
            document.MailMerge.ExecuteNestedGroup(GetRelationalData());
            // Save the result Word document.
            document.Save(Path.GetFullPath("../../../Output/Output.docx"));
            // Close the Word document
            document.Close();
        }
        #region Helper Method
        static MailMergeDataTable GetRelationalData()
        {            
            //Gets data from XML
            Stream xmlStream = File.OpenRead(Path.GetFullPath(@"Data/InvoiceDetails.xml"));
            XmlDocument xmlDocument = new XmlDocument();
            xmlDocument.Load(xmlStream);
            xmlStream.Dispose();

            ExpandoObject customerDetails = new ExpandoObject();
            GetDataAsExpandoObject((xmlDocument as XmlNode).LastChild, ref customerDetails);
            // Treat customerDetails as a dictionary
            IDictionary<string, object> customerDict = customerDetails as IDictionary<string, object>;
            // Get the "Invoices" list
            List<ExpandoObject> invoicesList = customerDict["Invoices"] as List<ExpandoObject>;
            // Take the first item in that list
            IDictionary<string, object> firstInvoiceGroup = invoicesList[0] as IDictionary<string, object>;
            // Get the "Invoice" list from that group
            List<ExpandoObject>  invoices = firstInvoiceGroup["Invoice"] as List<ExpandoObject>;            
            //Creates an instance of "MailMergeDataTable" by specifying mail merge group name and "IEnumerable" collection
            MailMergeDataTable dataTable = new MailMergeDataTable("Invoices", invoices);
            return dataTable;
        }
        /// <summary>
        /// Gets the data as ExpandoObject.
        /// </summary>
        /// <param name="reader">The reader.</param>
        /// <returns></returns>
        /// <exception cref="System.Exception">reader</exception>
        /// <exception cref="XmlException">Unexpected xml tag  + reader.LocalName</exception>
        private static void GetDataAsExpandoObject(XmlNode node, ref ExpandoObject dynamicObject)
        {
            if (node.InnerText == node.InnerXml)
                dynamicObject.TryAdd(node.LocalName, node.InnerText);
            else
            {
                List<ExpandoObject> childObjects;
                if ((dynamicObject as IDictionary<string, object>).ContainsKey(node.LocalName))
                    childObjects = (dynamicObject as IDictionary<string, object>)[node.LocalName] as List<ExpandoObject>;
                else
                {
                    childObjects = new List<ExpandoObject>();
                    dynamicObject.TryAdd(node.LocalName, childObjects);
                }
                ExpandoObject childObject = new ExpandoObject();
                foreach (XmlNode childNode in (node as XmlNode).ChildNodes)
                {
                    GetDataAsExpandoObject(childNode, ref childObject);
                }
                childObjects.Add(childObject);
            }
        }
        #endregion
    }
}

