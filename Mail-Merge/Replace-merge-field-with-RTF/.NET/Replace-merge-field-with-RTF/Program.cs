using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.Data;

namespace Replace_merge_field_with_RTF
{
    class Program
    {
        static void Main(string[] args)
        {
            using (FileStream fileStream = new FileStream(Path.GetFullPath(@"Data/Template.docx"), FileMode.Open, FileAccess.ReadWrite))
            {
                //Open an existing Word document.
                using (WordDocument document = new WordDocument(fileStream, FormatType.Docx))
                {
                    //Gets data to perform the mail merge.
                    DataTable table = GetDataTable();
                    //Executes the mail merge using the provided data.
                    document.MailMerge.Execute(table);
                    //Replaces placeholder text with content from RTF files.
                    InsertRTF(document);
                    //Creates file stream.
                    using (FileStream outputStream = new FileStream(Path.GetFullPath(@"Output/Result.docx"), FileMode.Create, FileAccess.ReadWrite))
                    {
                        //Saves the Word document to file stream.
                        document.Save(outputStream, FormatType.Docx);
                    }
                }
            }
        }
        #region Helper methods
        /// <summary>
        /// Create and return a DataTable with sample data for the mail merge.
        /// </summary>
        /// <returns>A DataTable with the required data for mail merge.</returns>
        private static DataTable GetDataTable()
        {
            //Create a new DataTable with columns matching the merge fields.
            DataTable dataTable = new DataTable("RTF");
            dataTable.Columns.Add("CustomerName");
            dataTable.Columns.Add("Address");
            dataTable.Columns.Add("Phone");
            dataTable.Columns.Add("ProductList");

            //Create a new row and add sample data.
            DataRow datarow = dataTable.NewRow();
            dataTable.Rows.Add(datarow);
            datarow["CustomerName"] = "Nancy Davolio";
            datarow["Address"] = "59 rue de I'Abbaye, Reims 51100, France";
            datarow["Phone"] = "1-888-936-8638";

            //Append placeholder text for RTF content into the row.
            datarow["ProductList"] = "#InsertRTF_ProductList#";
            return dataTable;
        }
        /// <summary>
        /// Replace the placeholders in the Word document with RTF content.
        /// </summary>
        /// <param name="document">The Word document where placeholders are replaced.</param>
        private static void InsertRTF(WordDocument document)
        {
            using (FileStream fileStream = new FileStream(Path.GetFullPath(@"Data/ProductList.rtf"), FileMode.Open, FileAccess.ReadWrite))
            {
                //Open the RTF document.
                using (WordDocument rtfDocument = new WordDocument(fileStream, FormatType.Rtf))
                {
                    //Replace the placeholder with RTF content.
                    document.Replace("#InsertRTF_ProductList#", rtfDocument, true, true);
                }
            }
        }
        #endregion
    }
}