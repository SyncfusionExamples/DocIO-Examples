using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
namespace Add_docvariable_in_hyperlink_field
{
    internal class Program
    {
        static void Main(string[] args)
        {
            // Creates a new instance of WordDocument (empty Word document).
            using (WordDocument wordDocument = new WordDocument())
            {
                // Adds a section to the document.
                IWSection section = wordDocument.AddSection();
                // Adds a paragraph to the section.
                IWParagraph paragraph = section.AddParagraph();
                // Adds a document variable named "Author" with a value of "https://www.syncfusion.com".
                wordDocument.Variables.Add("Author", "https://www.syncfusion.com");
                // Appends a hyperlink to the paragraph.
                WField hyperlinkField = paragraph.AppendHyperlink("https://www.google.com", "google", HyperlinkType.WebLink) as WField;
                // Embeds the document variable within the hyperlink field.
                AddDocVariableToHyperlink(paragraph, hyperlinkField);
                // Updates all document fields to reflect changes.
                wordDocument.UpdateDocumentFields();
                // Saves the output Word document to a specified file stream in DOCX format.
                using (FileStream stream = new FileStream(Path.GetFullPath(@"../../../Output/output.docx"), FileMode.OpenOrCreate))
                {
                    wordDocument.Save(stream, FormatType.Docx);
                }
            }
        }

        /// <summary>
        /// Modifies a hyperlink field to include a document variable.
        /// </summary>
        private static void AddDocVariableToHyperlink(IWParagraph paragraph, WField? field)
        {
            // Get the index of the field in the paragraph to insert the field code.
            int fieldIndex = paragraph.Items.IndexOf(field) + 1;
            // Add the field code for the hyperlink.
            field.FieldCode = "HYPERLINK \"";
            // Increment field index for further insertion.
            fieldIndex++;
            // Insert document variables into the field code at the specified index.
            InsertDocVariables("Author", ref fieldIndex, paragraph);
            // Close the hyperlink field with the ending quote.
            InsertText("\"", ref fieldIndex, paragraph as WParagraph);
        }

        /// <summary>
        /// Inserts text such as operators or identifiers into the paragraph.
        /// </summary>
        private static void InsertText(string text, ref int fieldIndex, WParagraph paragraph)
        {
            // Create a new text range with the specified text.
            WTextRange textRange = new WTextRange(paragraph.Document);
            textRange.Text = text;
            // Insert the text range as a field code item at the specified index.
            paragraph.Items.Insert(fieldIndex, textRange);
            // Increment the field index for subsequent insertions.
            fieldIndex++;
        }
        /// <summary>
        /// Inserts document variables at the given index in the specified paragraph.
        /// </summary>
        private static void InsertDocVariables(string fieldName, ref int fieldIndex, IWParagraph paragraph)
        {
            // Create a new paragraph to hold the document variable field.
            WParagraph para = new WParagraph(paragraph.Document);
            para.AppendField(fieldName, FieldType.FieldDocVariable);
            // Get the count of child entities in the paragraph (should be 1 for the field).
            int count = para.ChildEntities.Count;
            // Insert the field into the original paragraph, ensuring the complete field structure is added.
            paragraph.ChildEntities.Insert(fieldIndex, para.ChildEntities[0]);
            // Increment the field index based on the count of inserted entities.
            fieldIndex += count;
        }
    }
}
