using Syncfusion.DocIO.DLS;
using Syncfusion.DocIO;

namespace Replace_MergeField_with_Bookmark_Hyperlink
{
    internal class Program
    {
        static void Main(string[] args)
        {
            // Open the existing Word document using a FileStream.
            using (FileStream docStream = new FileStream(Path.GetFullPath(@"Data/Template.docx"), FileMode.Open, FileAccess.Read))
            {
                // Load the Word document into the WordDocument object.
                using (WordDocument document = new WordDocument(docStream, FormatType.Docx))
                {
                    // Find the first merge field with the name "Test" in the document.
                    WMergeField mergeField = document.FindItemByProperty(EntityType.MergeField, "FieldName", "Test") as WMergeField;

                    // Convert the merge field to a placeholder and get the placeholder text.
                    string textToFind = ConvertMergeFieldToPlaceHolder(mergeField);

                    // Replace the placeholder text with "2" in the document.
                    document.Replace(textToFind, "2", false, true);

                    // Create a new TextBodyPart to hold the bookmark hyperlink.
                    TextBodyPart textBodyPart = new TextBodyPart(document);

                    // Create a new paragraph to add the hyperlink.
                    WParagraph paragraphHyperlink = new WParagraph(document);
                    textBodyPart.BodyItems.Add(paragraphHyperlink);

                    // Append a hyperlink to the paragraph that links to the bookmark "Bookmark1" with display text "Syncfusion".
                    paragraphHyperlink.AppendHyperlink("Bookmark1", "Syncfusion", HyperlinkType.Bookmark);

                    // Replace the text "2" with the bookmark hyperlink in the document.
                    document.Replace("2", textBodyPart, false, true);

                    // Save the modified document to a new file.
                    using (FileStream outputStream = new FileStream(Path.GetFullPath(@"Output/Result.docx"), FileMode.Create, FileAccess.Write))
                    {
                        document.Save(outputStream, FormatType.Docx);
                    }
                }
            }
        }

        // This method converts a merge field to a placeholder by replacing the merge field with its text representation.
        private static string ConvertMergeFieldToPlaceHolder(WMergeField field)
        {
            // Get the paragraph that contains the merge field.
            WParagraph paragraph = field.OwnerParagraph;

            // Find the index of the merge field in the paragraph's child entities.
            int itemIndex = paragraph.ChildEntities.IndexOf(field);

            // Create a new WTextRange to hold the text that will replace the merge field.
            WTextRange textRange = new WTextRange(paragraph.Document);

            // Get the text from the merge field (which may be a hyperlink field).
            textRange.Text = GetMergeFieldText(itemIndex, paragraph);

            // Remove the merge field from the paragraph's child entities.
            paragraph.ChildEntities.RemoveAt(itemIndex);

            // Insert the text that replaces the merge field at the same position.
            paragraph.ChildEntities.Insert(itemIndex, textRange);

            // Return the text that will replace the merge field.
            return textRange.Text;
        }

        // This method extracts the text of a merge field, handling nested fields if present.
        private static string GetMergeFieldText(int hyperlinkIndex, WParagraph paragraph)
        {
            string text = string.Empty;

            // Stack to handle nested fields while iterating through paragraph child entities.
            Stack<Entity> fieldStack = new Stack<Entity>();
            fieldStack.Push(paragraph.ChildEntities[hyperlinkIndex]);

            // Flag to control whether to collect text or skip field code sections.
            bool isFieldCode = true;
            int i = (hyperlinkIndex + 1);

            // Iterate through the paragraph's items (child entities) to collect text between field separator and field end.
            while (i < paragraph.Items.Count)
            {
                Entity item = paragraph.ChildEntities[i];

                // Check if the current item is a field and handle it by adding to the stack.
                if (item is WField)
                {
                    fieldStack.Push(item);
                    isFieldCode = true;  // Set flag to skip text in field code section.
                }
                // Check if the current item is a field separator and enable collecting text.
                else if (item is WFieldMark && ((WFieldMark)item).Type == FieldMarkType.FieldSeparator)
                {
                    isFieldCode = false;
                }
                // Check if the current item is a field end and handle field stack.
                else if (item is WFieldMark && ((WFieldMark)item).Type == FieldMarkType.FieldEnd)
                {
                    // If it's the end of the outermost field, return the collected text.
                    if (fieldStack.Count == 1)
                    {
                        fieldStack.Clear();
                        return text;
                    }
                    else
                    {
                        fieldStack.Pop();
                    }
                }
                // If not in field code section and the item is a text range, collect the text.
                else if (!isFieldCode && item is WTextRange)
                {
                    text += ((WTextRange)item).Text;
                }

                i++;
            }

            return text;
        }
    }
}
