using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;

namespace Find_and_remove_table_by_title
{
    internal class Program
    {
        static void Main(string[] args)
        {
            using (FileStream fileStream = new FileStream(Path.GetFullPath(@"Data/Template.docx"), FileMode.Open, FileAccess.ReadWrite))
            {
                using (WordDocument document = new WordDocument(fileStream, FormatType.Automatic))
                {
                    // Find the table with title.
                    WTable table = document.FindItemByProperty(EntityType.Table, "Title", "Product Overview") as WTable;

                    if (table != null)
                    {
                        // Remove table from document
                        table.OwnerTextBody.ChildEntities.Remove(table);
                    }

                    using (FileStream outputStream = new FileStream(Path.GetFullPath(@"Output/Result.docx"), FileMode.Create, FileAccess.ReadWrite))
                    {
                        document.Save(outputStream, FormatType.Docx);
                    }
                }
            }
        }
    }
}
