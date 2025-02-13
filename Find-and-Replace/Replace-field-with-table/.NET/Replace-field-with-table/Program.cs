using Syncfusion.DocIO.DLS;
using Syncfusion.DocIO;

namespace Replace_field_with_table
{
    internal class Program
    {
        static void Main(string[] args)
        {
            // Open the existing Word document using a FileStream.
            using (FileStream docStream = new FileStream(Path.GetFullPath(@"Data/InputDocument.docx"), FileMode.Open, FileAccess.Read))
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

                    //Create a new table
                    IWTable table = new WTable(document);
                    table.ResetCells(3, 2);
                    table.Rows[0].Cells[0].AddParagraph().AppendText("Sno");
                    table.Rows[0].Cells[1].AddParagraph().AppendText("Product");
                    table.Rows[0].IsHeader = true;
                    table.Rows[1].Cells[0].AddParagraph().AppendText("1.");
                    table.Rows[1].Cells[1].AddParagraph().AppendText("Essential DocIO");
                    table.Rows[2].Cells[0].AddParagraph().AppendText("2.");
                    table.Rows[2].Cells[1].AddParagraph().AppendText("Essential Pdf");

                    // Clone the generated table.
                    IWTable table1 = table.Clone() as IWTable;
                    // Insert the cloned table after the paragraph containing the SEQ field.
                    paragraph.OwnerTextBody.ChildEntities.Insert(paraindex + 1, table1);
                    // Insert the modified cloned paragraph after the inserted table.
                    paragraph.OwnerTextBody.ChildEntities.Insert(paraindex + 2, clonedParagraph);
                    // Save the modified document to a new file.
                    using (FileStream docStream1 = new FileStream(Path.GetFullPath(@"Output/Result.docx"), FileMode.Create, FileAccess.Write))
                    {
                        document.Save(docStream1, FormatType.Docx);
                    }
                }
            }
        }
    }
}
