using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.Collections.Generic;
using System.IO;

namespace ExtractTablesFromWordDocument
{
    class Program
    {
        static void Main(string[] args)
        {
            // Open the template Word document.
            using (FileStream inputFileStream = new FileStream(Path.GetFullPath(@"Data/Template.docx"), FileMode.Open, FileAccess.ReadWrite))
            {
                using (WordDocument sourceDocument = new WordDocument(inputFileStream, FormatType.Docx))
                {
                    using (WordDocument newDocument = new WordDocument())
                    {
                        newDocument.EnsureMinimal();
                        // Find all tables in the Word document.
                        List<Entity> tableEntities = sourceDocument.FindAllItemsByProperty(EntityType.Table, null, null);

                        // Iterate through each table and add it to the new document.
                        foreach (WTable table in tableEntities)
                        {
                            WTable clonedTable = table.Clone() as WTable;

                            // Add the cloned table to the new document.
                            newDocument.LastSection.Body.ChildEntities.Add(clonedTable);

                            // Add an empty paragraph for spacing between tables.
                            newDocument.LastSection.Body.AddParagraph();
                        }

                        // Save the new document with extracted tables.
                        using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"Output/Result.docx"), FileMode.Create, FileAccess.ReadWrite))
                        {
                            newDocument.Save(outputFileStream, FormatType.Docx);
                        }
                    }
                }
            }
        }
    }
}
