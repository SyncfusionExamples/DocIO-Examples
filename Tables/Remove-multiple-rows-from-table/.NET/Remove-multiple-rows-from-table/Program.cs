using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;

namespace Remove_multiple_rows_from_table
{
    class Program
    {
        static void Main(string[] args)
        {
            // Open the template document from the specified file path.
            using (FileStream inputFileStream = new FileStream(Path.GetFullPath(@"Data/Template.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                // Load the Word document.
                using (WordDocument document = new WordDocument(inputFileStream, FormatType.Automatic))
                {
                    // Access the first section of the document.
                    WSection section = document.Sections[0];
                    // Access the first table in the section.
                    WTable table = section.Tables[0] as WTable;
                    // Iterate through the table rows in reverse order to prevent index shifting issues.
                    for (int i = table.Rows.Count - 1; i >= 0; i--)
                    {
                        // Remove specific rows based on index (e.g., rows at index 2, 5, and 7).
                        if (i == 2 || i == 5 || i == 7)
                        {
                            WTableRow row = table.Rows[i];
                            table.Rows.Remove(row);
                        }
                    }
                    // Save the updated document to the output file.
                    using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"Output/Result.docx"), FileMode.Create, FileAccess.ReadWrite))
                    {
                        document.Save(outputFileStream, FormatType.Docx);
                    }
                }
            }
        }
    }
}
