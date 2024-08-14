using Syncfusion.DocIO.DLS;
using Syncfusion.DocIO;

namespace Replace_field_with_table
{
    internal class Program
    {
        static void Main(string[] args)
        {
            // Open the existing Word document using a FileStream.
            using (FileStream docStream = new FileStream(Path.GetFullPath(@"../../../Data/InputDocument.docx"), FileMode.Open, FileAccess.Read))
            {
                // Load the existing Word document.
                using (WordDocument document = new WordDocument(docStream, FormatType.Docx))
                {
                    // Find the first sequence field (SEQ field) in the document.
                    WSeqField seqField = document.FindItemByProperty(EntityType.SeqField, "", "") as WSeqField;

                    // Get the paragraph that contains the SEQ field.
                    WParagraph paragraph = seqField.OwnerParagraph;

                    // Get the index of the paragraph within the text body.
                    int paraindex = paragraph.OwnerTextBody.ChildEntities.IndexOf(paragraph);

                    // Get the index of the SEQ field within the paragraph.
                    int seqfieldIndex = paragraph.ChildEntities.IndexOf(seqField);

                    // Clone the paragraph that contains the SEQ field.
                    WParagraph clonedParagraph = seqField.OwnerParagraph.Clone() as WParagraph;

                    // Remove all entities before the SEQ field index in the cloned paragraph.
                    for (int i = seqfieldIndex; i >= 0; i--)
                    {
                        clonedParagraph.ChildEntities.RemoveAt(i);
                    }

                    // Remove all entities from the SEQ field index onward in the original paragraph.
                    for (int j = paragraph.ChildEntities.Count - 1; j >= seqfieldIndex; j--)
                    {
                        paragraph.ChildEntities.RemoveAt(j);
                    }

                    // Generate a sample table.
                    IWTable table = GetTable();

                    // Clone the generated table.
                    IWTable table1 = table.Clone() as IWTable;

                    // Insert the cloned table after the paragraph containing the SEQ field.
                    paragraph.OwnerTextBody.ChildEntities.Insert(paraindex + 1, table1);

                    // Insert the modified cloned paragraph after the inserted table.
                    paragraph.OwnerTextBody.ChildEntities.Insert(paraindex + 2, clonedParagraph);

                    // Save the modified document to a new file.
                    using (FileStream docStream1 = new FileStream(Path.GetFullPath(@"../../../Data/ResultDocument.docx"), FileMode.Create, FileAccess.Write))
                    {
                        document.Save(docStream1, FormatType.Docx);
                    }
                }
            }
        }

        // Method to generate a sample table.
        static IWTable GetTable()
        {
            // Creates a new Word document.
            WordDocument document = new WordDocument();

            // Adds a section into the Word document.
            IWSection section = document.AddSection();

            // Adds a paragraph with the text "Price Details" in bold and Arial font.
            IWTextRange textRange = section.AddParagraph().AppendText("Price Details");
            textRange.CharacterFormat.FontName = "Arial";
            textRange.CharacterFormat.FontSize = 12;
            textRange.CharacterFormat.Bold = true;

            // Adds an empty paragraph (for spacing).
            section.AddParagraph();

            // Adds a new table with 3 rows and 2 columns.
            IWTable table = section.AddTable();
            table.ResetCells(3, 2);

            // Adds the column headers to the first row.
            textRange = table[0, 0].AddParagraph().AppendText("Item");
            textRange.CharacterFormat.FontName = "Arial";
            textRange.CharacterFormat.FontSize = 12;
            textRange.CharacterFormat.Bold = true;

            textRange = table[0, 1].AddParagraph().AppendText("Price($)");
            textRange.CharacterFormat.FontName = "Arial";
            textRange.CharacterFormat.FontSize = 12;
            textRange.CharacterFormat.Bold = true;

            // Adds the first item and its price to the second row.
            textRange = table[1, 0].AddParagraph().AppendText("Cycle 1");
            textRange.CharacterFormat.FontName = "Arial";
            textRange.CharacterFormat.FontSize = 10;

            textRange = table[1, 1].AddParagraph().AppendText("500");
            textRange.CharacterFormat.FontName = "Arial";
            textRange.CharacterFormat.FontSize = 10;

            // Adds the second item and its price to the third row.
            textRange = table[2, 0].AddParagraph().AppendText("Cycle 2");
            textRange.CharacterFormat.FontName = "Arial";
            textRange.CharacterFormat.FontSize = 10;

            textRange = table[2, 1].AddParagraph().AppendText("300");
            textRange.CharacterFormat.FontName = "Arial";
            textRange.CharacterFormat.FontSize = 10;

            // Returns the generated table.
            return table;
        }
    }
}
