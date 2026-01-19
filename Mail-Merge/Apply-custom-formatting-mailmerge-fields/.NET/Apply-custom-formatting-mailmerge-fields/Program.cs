using Syncfusion.DocIO.DLS;
using System.Data;

namespace Apply_custom_formatting_mailmerge_fields
{
    class Program
    {
        public static Dictionary<WParagraph, List<KeyValuePair<int, string>>> paratoModifyNumberFormat
           = new Dictionary<WParagraph, List<KeyValuePair<int, string>>>();
        public static Dictionary<WParagraph, List<KeyValuePair<int, string>>> paratoModifyDateFormat
            = new Dictionary<WParagraph, List<KeyValuePair<int, string>>>();
        public static void Main(string[] args)
        {
            // Load the existing word document
            WordDocument document = new WordDocument(Path.GetFullPath(@"Data\Template.docx"));
            // Enable separate page for each invoice
            document.MailMerge.StartAtNewPage = true;
            document.MailMerge.MergeField += new MergeFieldEventHandler(MergeFieldEvent);
            // Perform mail merge
            document.MailMerge.ExecuteGroup(GetInvoiceData());
            // Update the merge field results with formatted values.
            UpdateMergeFieldResult(true);
            UpdateMergeFieldResult(false);
            // Save the Word document.
            document.Save(Path.GetFullPath(@"../../../Output/Output.docx"));
            // Close the document
            document.Close();

        }
        /// <summary>
        /// Event handler triggered during mail merge for each merge field.
        /// </summary>
        /// <param name="sender">The source of the event (MailMerge engine)</param>
        /// <param name="args">Provides information about the current merge field.</param>
        private static void MergeFieldEvent(object sender, MergeFieldEventArgs args)
        {
            // Get the mergefield's Owner paragraph
            WParagraph mergeFieldOwnerParagraph = args.CurrentMergeField.OwnerParagraph;
            // Find the index of the current merge field within the paragraph.
            int index = mergeFieldOwnerParagraph.ChildEntities.IndexOf(args.CurrentMergeField);
            if (args.FieldName == "Amount")
            {
                // Check if this paragraph already has an entry in the dictionary.
                // If not, create a new list to store field index and field value.
                if (!paratoModifyNumberFormat.TryGetValue(mergeFieldOwnerParagraph, out var fields))
                {
                    fields = new List<KeyValuePair<int, string>>();
                    paratoModifyNumberFormat[mergeFieldOwnerParagraph] = fields;
                }
                // Add the current merge field's index and field name
                fields.Add(new KeyValuePair<int, string>(index, args.FieldValue.ToString()));
            }
            else if (args.FieldName == "InvoiceDate")
            {
                // Check if this paragraph already has an entry in the dictionary.
                // If not, create a new list to store field index and field value.
                if (!paratoModifyDateFormat.TryGetValue(mergeFieldOwnerParagraph, out var fields))
                {
                    fields = new List<KeyValuePair<int, string>>();
                    paratoModifyDateFormat[mergeFieldOwnerParagraph] = fields;
                }
                // Add the current merge field's index and field name
                fields.Add(new KeyValuePair<int, string>(index, args.FieldValue.ToString()));
            }
        }
        /// <summary>
        /// Updates the merge fields result after mail merge by applying number and date formatting using IF fields.
        /// </summary>       
        /// <param name="numberType">The boolean denotes current changes as Number format</param>
        public static void UpdateMergeFieldResult(bool numberType)
        {
            Dictionary<WParagraph, List<KeyValuePair<int, string>>> tempDictonary;
            if (numberType)
                tempDictonary = paratoModifyNumberFormat;
            else
                tempDictonary = paratoModifyDateFormat;
            // Iterate the outer dictionary entries
            foreach (var dictionaryItem in tempDictonary)
            {
                // Get the merge field result paragraph
                WParagraph mergeFieldParagraph = dictionaryItem.Key;
                // The list of (index, fieldValues) pairs for this paragraph.
                var fieldList = dictionaryItem.Value;                
                for (int i = 0; i <= fieldList.Count - 1; i++)
                {
                    // Get the index and Field values ("Number" or "Date")
                    int index = fieldList[i].Key;
                    string fieldValue = fieldList[i].Value;
                    // Get the existing merge field result text at the specified index.
                    WTextRange mergeFieldText = (WTextRange)mergeFieldParagraph.ChildEntities[index];
                    if (mergeFieldText != null)
                    {
                        // Create the temporary document and insert the IF field.
                        WordDocument tempDocument = new WordDocument();
                        WSection section = (WSection)tempDocument.AddSection();
                        WParagraph ifFieldParagraph = (WParagraph)section.AddParagraph();
                        WIfField field = (WIfField)ifFieldParagraph.AppendField("IfField", Syncfusion.DocIO.FieldType.FieldIf);
                        // Check if the Number field value
                        if (numberType)
                        {
                            // Format number: 1,234.56 
                            field.FieldCode = $"IF 1 = 1 \"{fieldValue}\" \" \" \\# \"#,##0.00";
                        }
                        // Update the Date field value
                        else
                        {
                            // Format date: dd/MMM/yyyy
                            field.FieldCode = $"IF 1 = 1 \"{fieldValue}\" \" \" \\@ \"dd/MMM/yyyy\" ";
                        }
                        // Update the field and unlink
                        tempDocument.UpdateDocumentFields();
                        field.Unlink();
                        // Update the Merge field result
                        WTextRange modifiedText = (WTextRange)ifFieldParagraph.ChildEntities[0];
                        mergeFieldText.Text = modifiedText.Text;
                        // close the temp document
                        tempDocument.Close();
                    }
                }
            }
            if (numberType)
                paratoModifyNumberFormat.Clear();
            else
                paratoModifyDateFormat.Clear();
        }
        private static DataTable GetInvoiceData()
        {
            DataTable table = new DataTable("Invoice");

            table.Columns.Add("InvoiceNumber");
            table.Columns.Add("InvoiceDate");
            table.Columns.Add("CustomerName");
            table.Columns.Add("ItemDescription");
            table.Columns.Add("Amount");           
            // First Invoice
            table.Rows.Add("INV001", "2024-05-01", "Andy Bernard", "Consulting Services", "3,000.578");
            // Second Invoice
            table.Rows.Add("INV002", "2024-05-05", "Stanley Hudson", "Software Development", "4,500.052");
            // Third Invoice
            table.Rows.Add("INV003", "2024-05-10", "Margaret Peacock", "UI Design Services", "2,000.600");

            return table;
        }
    }
}