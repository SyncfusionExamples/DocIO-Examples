using Syncfusion.DocIO.DLS;
using Syncfusion.DocIO;

namespace Add_page_break_between_rows
{
    internal class Program
    {
        static void Main(string[] args)
        {
            // Load the Word document from the specified file path.
            using (FileStream inputStream = new FileStream(Path.GetFullPath(@"Data/Template.docx"), FileMode.Open, FileAccess.Read))
            {
                using (WordDocument document = new WordDocument(inputStream, FormatType.Docx))
                {
                    // Access the table from the specified index in the document's body.
                    WTable table = document.Sections[0].Body.ChildEntities[4] as WTable;

                        // Check if the table contains 2 or more rows.
                        if (table != null && table.Rows.Count >= 2)
                        {
                            // Clone and remove the first row of the table.
                            WTableRow clonedRow = table.Rows[0].Clone();
                            table.Rows.RemoveAt(0);

                            // Get the text body of the table.
                            WTextBody documentBody = table.OwnerTextBody;

                            // Determine the index of the current table in the document body.
                            int currentTableIndex = documentBody.ChildEntities.IndexOf(table);

                            // Create a new paragraph and add a page break.
                            WParagraph pageBreakParagraph = new WParagraph(document);
                            pageBreakParagraph.AppendBreak(BreakType.PageBreak);

                            // Insert the new paragraph (with page break) before the current table.
                            documentBody.ChildEntities.Insert(currentTableIndex, pageBreakParagraph);

                            // Create a new table and insert it before the page break paragraph.
                            WTable newTable = new WTable(document);
                            documentBody.ChildEntities.Insert(currentTableIndex, newTable);

                            // Add the cloned row to the newly created table.
                            newTable.Rows.Add(clonedRow);
                        }
                    // Save the modified document to the specified file path.
                    using (FileStream outputStream = new FileStream(Path.GetFullPath(@"Output/Result.docx"), FileMode.Create, FileAccess.Write))
                    {
                        document.Save(outputStream, FormatType.Docx);
                    }
                }
            }
        }
    }
}

