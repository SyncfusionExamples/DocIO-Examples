using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.Drawing;
using System.IO;

namespace Apply_shading_to_fields
{
    class Program
    {
        static void Main(string[] args)
        {
            // Creates a new instance of WordDocument to work with.
            using (WordDocument document = new WordDocument())
            {
                // Add a new section to the Word document.
                IWSection section = document.AddSection();
                IWParagraph paragraph = section.AddParagraph();

                // Add text before the field (e.g., "Date Field - ").
                IWTextRange firstText = paragraph.AppendText("Date Field - ");

                // Adds a new Date field to the paragraph with the specified format.
                WField field = paragraph.AppendField("Date", FieldType.FieldDate) as WField;
                // Set the field code to display the date in "MMMM d, yyyy" format.
                field.FieldCode = @"DATE  \@" + "\"MMMM d, yyyy\"";

                // Reference the field as an entity.
                IEntity entity = field;

                // Apply shading to the field.
                ApplyShading(entity);

                // Add another paragraph for the "If" field.
                paragraph = section.AddParagraph();
                paragraph.AppendText("If Field - ");

                // Creates a new IF field.
                field = paragraph.AppendField("If", FieldType.FieldIf) as WIfField;
                // Specifies the field code for the IF statement with true and false branches.
                field.FieldCode = "IF \"True\" = \"True\" \"The given statement is Correct\" \"The given statement is Wrong\"";
                entity = field;

                // Apply shading to the field.
                ApplyShading(entity);

                // Updates the fields in the document.
                document.UpdateDocumentFields();

                // Create a file stream to save the document.
                using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"Output/Result.docx"), FileMode.Create, FileAccess.ReadWrite))
                {
                    // Saves the Word document to the specified file stream in DOCX format.
                    document.Save(outputFileStream, FormatType.Docx);
                }
            }
        }

        // Method to apply shading (highlight) to the content of a field.
        static void ApplyShading(IEntity entity)
        {
            // Loop through the siblings of the current entity (field and its contents) until reaching the FieldEnd.
            while (entity.NextSibling != null)
            {
                // Check if the entity is a text range.
                if (entity is WTextRange)
                {
                    // Set the highlight color to LightGray for the text range.
                    (entity as WTextRange).CharacterFormat.HighlightColor = Color.LightGray;
                }
                // Check if the entity is a field mark and is of FieldEnd type.
                else if ((entity is WFieldMark) && (entity as WFieldMark).Type == FieldMarkType.FieldEnd)
                {
                    // Break out of the loop once the end of the field is reached.
                    break;
                }

                // Move to the next sibling entity (next part of the field).
                entity = entity.NextSibling;
            }
        }
    }
}
