using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;

namespace Split_table_with_same_format
{
    class Program
    {
        static void Main(string[] args)
        {
            //Open the file as a stream.
            using (FileStream inputStream = new FileStream(Path.GetFullPath(@"../../../Input.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                //Load the file stream into a Word document.
                using (WordDocument document = new WordDocument(inputStream, FormatType.Docx))
                {
                    //Access table in a Word document.
                    WTable table = document.Sections[0].Tables[0] as WTable;
                    //Clone the table.
                    WTable clonedTable = table.Clone();
                    int rowIndex = 2;
                    SplitTable(table, clonedTable, rowIndex);
                    //Add the second table.
                    table.OwnerTextBody.ChildEntities.Add(clonedTable);
                    //Create a file stream.
                    using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"../../../Sample.docx"), FileMode.Create, FileAccess.ReadWrite))
                    {
                        //Save the Word document to the file stream.
                        document.Save(outputFileStream, FormatType.Docx);
                    }
                }
            }
        }
        /// <summary>
        /// Split the table depending on the row index.
        /// </summary>
        private static void SplitTable(WTable table, WTable clonedTable, int rowIndex)
        {
            //Remove rows from the table.
            while (rowIndex < table.Rows.Count)
            {
                table.Rows.Remove(table.Rows[rowIndex]);
            }
            while (rowIndex != 0)
            {
                clonedTable.Rows.Remove(clonedTable.Rows[0]);
                rowIndex--;
            }
        }
    }
}
