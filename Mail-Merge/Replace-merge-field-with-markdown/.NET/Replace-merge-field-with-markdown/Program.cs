using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Net;
using System.Text;

namespace ReplaceMergeFieldWithMarkdown
{
    class Program
    {
        // Dictionary to store paragraphs and corresponding merge field positions with Markdown content.
        static Dictionary<WParagraph, Dictionary<int, string>> paraToInsertMarkdown = new Dictionary<WParagraph, Dictionary<int, string>>();

        static void Main(string[] args)
        {
            // Open the template document.
            using (FileStream inputFileStream = new FileStream(Path.GetFullPath(@"Data/Template.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                using (WordDocument document = new WordDocument(inputFileStream, FormatType.Automatic))
                {
                    // Attach mail merge event handler to replace merge field with Markdown.
                    document.MailMerge.MergeField += MergeFieldEvent;

                    // Retrieve data for mail merge.
                    DataTable table = GetDataTable();

                    // Perform mail merge.
                    document.MailMerge.Execute(table);

                    // Insert the Markdown content at the specified positions.
                    InsertMarkdown();

                    // Detach mail merge event handler.
                    document.MailMerge.MergeField -= MergeFieldEvent;

                    // Save the updated document.
                    using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"Output/Result.docx"), FileMode.Create, FileAccess.ReadWrite))
                    {
                        document.Save(outputFileStream, FormatType.Docx);
                    }

                    // Close the document.
                    document.Close();
                }
            }
        }

        #region Helper Methods

        /// <summary>
        /// Event handler to replace merge fields with Markdown content.
        /// </summary>
        private static void MergeFieldEvent(object sender, MergeFieldEventArgs args)
        {
            if (args.TableName == "Markdown" && args.FieldName == "ProductList")
            {
                // Get the paragraph containing the merge field.
                WParagraph paragraph = args.CurrentMergeField.OwnerParagraph;

                // Get the merge field position in the paragraph.
                int mergeFieldIndex = paragraph.ChildEntities.IndexOf(args.CurrentMergeField);

                // Store the Markdown content along with its position in the paragraph.
                if (!paraToInsertMarkdown.ContainsKey(paragraph))
                {
                    paraToInsertMarkdown[paragraph] = new Dictionary<int, string>();
                }
                paraToInsertMarkdown[paragraph][mergeFieldIndex] = args.FieldValue.ToString();

                // Set the field value as empty to remove the original merge field.
                args.Text = string.Empty;
            }
        }

        /// <summary>
        /// Generates the data required for the mail merge operation.
        /// </summary>
        private static DataTable GetDataTable()
        {
            DataTable dataTable = new DataTable("Markdown");
            dataTable.Columns.Add("CustomerName");
            dataTable.Columns.Add("Address");
            dataTable.Columns.Add("Phone");
            dataTable.Columns.Add("ProductList");

            DataRow dataRow = dataTable.NewRow();
            dataRow["CustomerName"] = "Nancy Davolio";
            dataRow["Address"] = "59 rue de I'Abbaye, Reims 51100, France";
            dataRow["Phone"] = "1-888-936-8638";

            // Markdown content.
            dataRow["ProductList"] = "# Hello Markdown!\nThis is some **bold** text.";

            dataTable.Rows.Add(dataRow);
            return dataTable;
        }

        /// <summary>
        /// Inserts Markdown content at the correct positions in the Word document.
        /// </summary>
        private static void InsertMarkdown()
        {
            foreach (var paragraphEntry in paraToInsertMarkdown)
            {
                WParagraph paragraph = paragraphEntry.Key;
                Dictionary<int, string> fieldValues = paragraphEntry.Value;

                foreach (var fieldValueEntry in fieldValues)
                {
                    int index = fieldValueEntry.Key;
                    string markdownContent = fieldValueEntry.Value;

                    // Convert Markdown content to bytes and create a stream.
                    byte[] contentBytes = Encoding.UTF8.GetBytes(markdownContent);
                    using (MemoryStream memoryStream = new MemoryStream(contentBytes))
                    {
                        using (WordDocument markdownDoc = new WordDocument(memoryStream, FormatType.Markdown))
                        {
                            // Prepare to insert the Markdown content at the correct location.
                            TextBodyPart bodyPart = new TextBodyPart(paragraph.OwnerTextBody.Document);
                            BodyItemCollection bodyItems = bodyPart.BodyItems;

                            // Copy and paste the markdown at the same position of mergefield in Word document.
                            foreach (Entity entity in markdownDoc.LastSection.Body.ChildEntities)
                            {
                                bodyItems.Add(entity.Clone());
                            }
                            bodyPart.PasteAt(paragraph.OwnerTextBody, paragraph.OwnerTextBody.ChildEntities.IndexOf(paragraph), index);
                        }
                    }
                }
            }

            // Clear the dictionary after processing.
            paraToInsertMarkdown.Clear();
        }

        #endregion
    }
}
