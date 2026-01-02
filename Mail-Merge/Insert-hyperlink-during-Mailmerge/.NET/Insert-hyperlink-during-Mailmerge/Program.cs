using Syncfusion.DocIO.DLS;

namespace Insert_hyperlink_during_mailmerge
{
    class Program
    {
        public static void Main(string[] args)
        {
            // Open the template Word document
            WordDocument document = new WordDocument(Path.GetFullPath(@"Data/Template.docx"));
            // Attach the event handler that runs when a merge field is processed
            document.MailMerge.MergeField += MailMerge_MergeField;
            // Define the merge field names present in the template document
            string[] fieldNames = new string[] { "EmployeeId", "Name", "Phone", "City", "Contact" };
            // Define the values that will replace the merge fields during mail merge
            string[] fieldValues = new string[] { "1001", "Peter", "+122-2222222", "London", "peter@xyz.com" };
            //Execute Mail merge
            document.MailMerge.Execute(fieldNames, fieldValues);
            // Save the result document
            document.Save(Path.GetFullPath(@"../../../Output/output.docx"));
            // Close the Word document
            document.Close();
        }
        /// <summary>
        /// Event handler that customizes how merge fields are processed.
        /// </summary>
        /// <param name="sender">The source object raising the event (MailMerge engine).</param>
        /// <param name="args">Provides details about the current merge field being processed</param>
        private static void MailMerge_MergeField(object sender, MergeFieldEventArgs args)
        {

            // Check if the current merge field is "Contact", If Yes this field will be replaced with a hyperlink
            if (args.FieldName == "Contact")
            {
                // Create a new paragraph and append hyperlink, 
                WParagraph paragraph = new WParagraph(args.Document);
                WField hyperlink = paragraph.AppendHyperlink(args.FieldValue.ToString(), "Click ME", HyperlinkType.WebLink) as WField;
                // Get the current merge field object being processed
                WField mergeField = args.CurrentMergeField as WField;
                // Ensure the merge field exists before replacing it
                if (mergeField != null)
                {
                    // Get the paragraph that contains the merge field
                    WParagraph ownerParagraph = mergeField.OwnerParagraph;
                    // Insert the child entity (e.g., hyperlink) from the new paragraph into the original paragraph
                    for (int i = 0; i < paragraph.ChildEntities.Count; i++)
                    {
                        int fieldIndex = ownerParagraph.ChildEntities.IndexOf(mergeField);
                        ownerParagraph.ChildEntities.Insert(fieldIndex, paragraph.ChildEntities[i].Clone());
                    }
                    // Remove the original merge field from the paragraph
                    ownerParagraph.ChildEntities.Remove(mergeField);
                }
            }
        }
    }
}
