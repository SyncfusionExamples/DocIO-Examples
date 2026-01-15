using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.Collections.Generic;
using System.Data;
using System.IO;

namespace Replace_merge_field_with_HTML
{
    class Program
    {
        static Dictionary<WParagraph, List<KeyValuePair<int, string>>> paraToInsertHTML = new Dictionary<WParagraph, List<KeyValuePair<int, string>>>();
        static void Main(string[] args)
        {
            // Open the template Word document.
            using (FileStream fileStream = new FileStream(Path.GetFullPath(@"Data/Template.docx"), FileMode.Open, FileAccess.ReadWrite))
            {
                using (WordDocument document = new WordDocument(fileStream, FormatType.Docx))
                {
                    //Creates mail merge events handler to replace merge field with HTML.
                    document.MailMerge.MergeField += new MergeFieldEventHandler(MergeFieldEvent);
                    //Gets data to perform mail merge.
                    DataTable table = GetDataTable();
                    //Performs the mail merge.
                    document.MailMerge.Execute(table);
                    //Append HTML to paragraph.
                    InsertHtml();
                    //Removes mail merge events handler.
                    document.MailMerge.MergeField -= new MergeFieldEventHandler(MergeFieldEvent);
                    // Save the modified document.
                    using (FileStream outputStream = new FileStream(Path.GetFullPath(@"Output/Result.docx"), FileMode.Create, FileAccess.ReadWrite))
                    {
                        document.Save(outputStream, FormatType.Docx);
                    }
                }
            }
        }

        #region Helper methods
        /// <summary>
        /// Replaces merge field with HTML string by using MergeFieldEventHandler.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="args"></param>
        public static void MergeFieldEvent(object sender, MergeFieldEventArgs args)
        {
            if (args.TableName.Equals("HTML"))
            {
                if (args.FieldName.Equals("ProductList"))
                {
                    //Gets the current merge field owner paragraph.
                    WParagraph paragraph = args.CurrentMergeField.OwnerParagraph;
                    //Gets the current merge field index in the current paragraph.
                    int mergeFieldIndex = paragraph.ChildEntities.IndexOf(args.CurrentMergeField);
                    // Check if this paragraph already has an entry in the dictionary.
                    // If not, create a new list to store field index and value pairs.
                    if (!paraToInsertHTML.TryGetValue(paragraph, out var fields))
                    {
                        fields = new List<KeyValuePair<int, string>>();
                        paraToInsertHTML[paragraph] = fields;
                    }
                    // Add the current merge field's index and its value to the list
                    fields.Add(new KeyValuePair<int, string>(mergeFieldIndex, args.FieldValue.ToString()));
                    //Set field value as empty.
                    args.Text = string.Empty;
                }
            }
        }
        /// <summary>
        /// Gets the data to perform mail merge
        /// </summary>
        /// <returns></returns>
        private static DataTable GetDataTable()
        {
            DataTable dataTable = new DataTable("HTML");
            dataTable.Columns.Add("CustomerName");
            dataTable.Columns.Add("Address");
            dataTable.Columns.Add("Phone");
            dataTable.Columns.Add("ProductList");
            DataRow datarow = dataTable.NewRow();
            dataTable.Rows.Add(datarow);
            datarow["CustomerName"] = "Nancy Davolio";
            datarow["Address"] = "59 rue de I'Abbaye, Reims 51100, France";
            datarow["Phone"] = "1-888-936-8638";
            //Reads HTML string from the file.
            string htmlString = File.ReadAllText(Path.GetFullPath(@"Data/File.html"));
            datarow["ProductList"] = htmlString;
            return dataTable;
        }
        /// <summary>
        /// Append HTML to paragraph.
        /// </summary>
        private static void InsertHtml()
        {
            //Iterates through each item in the dictionary.
            foreach (KeyValuePair<WParagraph, List<KeyValuePair<int, string>>> dictionaryItems in paraToInsertHTML)
            {
                // Get the paragraph where HTML needs to be inserted.
                WParagraph paragraph = dictionaryItems.Key;
                // Get the list of (index, HTML string) pairs for this paragraph.
                List<KeyValuePair<int, string>> values = dictionaryItems.Value;
                // Iterate through the list in reverse order
                for (int i = values.Count - 1; i >= 0; i--)
                {
                    // Get the index of the merge field within the paragraph.
                    int index = values[i].Key;
                    // Get the HTML content to insert.
                    string fieldValue = values[i].Value;
                    // Get paragraph position.
                    int paragraphIndex = paragraph.OwnerTextBody.ChildEntities.IndexOf(paragraph);
                    //Inserts HTML string at the same position of mergefield in Word document.
                    paragraph.OwnerTextBody.InsertXHTML(fieldValue, paragraphIndex, index);
                }

            }
            paraToInsertHTML.Clear();
        }
        #endregion
    }
}
