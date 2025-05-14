using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.Dynamic;
using System.Xml;


// Dictionary to temporarily hold merge field HTML values and their positions.
Dictionary<WParagraph, Dictionary<int, string>> paraToInsertHTML = new Dictionary<WParagraph, Dictionary<int, string>>();

// Load the Word document template
using (WordDocument document = new WordDocument(Path.GetFullPath(@"Data/Template.docx")))
{
    // Load the XML file that contains data to be merged
    Stream xmlStream = System.IO.File.OpenRead(Path.GetFullPath(@"Data/Data.xml"));
    XmlDocument xmlDocument = new XmlDocument();
    xmlDocument.Load(xmlStream);
    xmlStream.Dispose(); // Close and release the file stream

    // Convert the XML data into a dynamic ExpandoObject structure
    ExpandoObject allDataObject = new ExpandoObject();
    GetDataAsExpandoObject(xmlDocument.LastChild, ref allDataObject);

    // Traverse the dynamic data to extract and map mail merge records
    if (allDataObject is IDictionary<string, object> allObjects &&
        allObjects.ContainsKey("Root"))
    {
        var rootObjects = (allObjects["Root"] as List<ExpandoObject>)[0] as IDictionary<string, object>;
        if (rootObjects.ContainsKey("Validator"))
        {
            var validatorObjects = (rootObjects["Validator"] as List<ExpandoObject>)[0] as IDictionary<string, object>;

            // Iterate over each component inside Validator (e.g., Login, Dashboard)
            foreach (var validatorObject in validatorObjects)
            {
                var componentMainTag = validatorObject.Key; // Typically "Component"
                var componentData = validatorObject.Value as List<ExpandoObject>;

                if (componentData != null)
                {
                    // Create a mail merge data table for each component
                    MailMergeDataTable componentDataTable = new MailMergeDataTable(componentMainTag, componentData);

                    // Attach event handler to handle HTML merge fields
                    document.MailMerge.MergeField += new MergeFieldEventHandler(MergeFieldEvent);

                    // Execute group mail merge using nested data table
                    document.MailMerge.ExecuteGroup(componentDataTable);
                }
            }
        }
    }

    // After mail merge is complete, insert HTML into placeholders
    InsertHtml();

    // Detach mail merge event handlers
    document.MailMerge.MergeField -= new MergeFieldEventHandler(MergeFieldEvent);

    // Save the updated document
    document.Save(Path.GetFullPath(@"Output/Result.docx"), FormatType.Docx);
}

/// <summary>
/// Event handler that intercepts the mail merge process to handle HTML content.
/// </summary>
void MergeFieldEvent(object sender, MergeFieldEventArgs args)
{
    if (args.FieldName.Equals("Description"))
    {
        // Get the paragraph containing the current merge field
        WParagraph paragraph = args.CurrentMergeField.OwnerParagraph;

        // Get the index of the merge field inside the paragraph
        int mergeFieldIndex = paragraph.ChildEntities.IndexOf(args.CurrentMergeField);

        // Store the HTML field value with its position
        Dictionary<int, string> fieldValues = new Dictionary<int, string>();
        fieldValues.Add(mergeFieldIndex, args.FieldValue.ToString());

        // Associate the paragraph with its corresponding HTML values
        paraToInsertHTML.Add(paragraph, fieldValues);

        // Remove the default text replacement; HTML will be inserted later
        args.Text = string.Empty;
    }
}

/// <summary>
/// Recursively parses XML nodes into ExpandoObject structure for easy access.
/// Supports handling of HTML content inside text nodes.
/// </summary>
void GetDataAsExpandoObject(XmlNode node, ref ExpandoObject dynamicObject)
{
    try
    {
        // Check if node is simple text (possibly with HTML tags)
        if (node.InnerText == node.InnerXml || (node.ChildNodes.Count == 1 && node.FirstChild.NodeType == XmlNodeType.Text))
        {
            if (!(dynamicObject as IDictionary<string, object>).ContainsKey(node.LocalName))
                (dynamicObject as IDictionary<string, object>).Add(node.LocalName, node.InnerText);
        }
        else
        {
            List<ExpandoObject> childObjects;

            // Reuse existing list if it already exists for this node
            if ((dynamicObject as IDictionary<string, object>).ContainsKey(node.LocalName))
                childObjects = (dynamicObject as IDictionary<string, object>)[node.LocalName] as List<ExpandoObject>;
            else
            {
                childObjects = new List<ExpandoObject>();
                (dynamicObject as IDictionary<string, object>).Add(node.LocalName, childObjects);
            }

            ExpandoObject childObject = new ExpandoObject();

            // Recursively parse each child node
            foreach (XmlNode childNode in node.ChildNodes)
            {
                GetDataAsExpandoObject(childNode, ref childObject);
            }

            childObjects.Add(childObject);
        }
    }
    catch (Exception e)
    {
        Console.WriteLine("Error in XML reading: " + e.ToString());
    }
}

/// <summary>
/// Replaces merge fields with actual HTML content using collected placeholders.
/// </summary>
void InsertHtml()
{
    // Iterate each paragraph and its associated HTML values
    foreach (KeyValuePair<WParagraph, Dictionary<int, string>> dictionaryItems in paraToInsertHTML)
    {
        WParagraph paragraph = dictionaryItems.Key;
        Dictionary<int, string> values = dictionaryItems.Value;

        foreach (KeyValuePair<int, string> valuePair in values)
        {
            int index = valuePair.Key;
            string fieldValue = valuePair.Value;

            // Insert HTML at the specified location in the paragraph
            paragraph.OwnerTextBody.InsertXHTML(fieldValue, paragraph.OwnerTextBody.ChildEntities.IndexOf(paragraph), index);
        }
    }

    // Clear dictionary after inserting all HTML
    paraToInsertHTML.Clear();
}
