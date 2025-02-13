using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Text;

namespace Merge_markdown_during_mail_merge
{
    class Program
    {
        static Dictionary<WParagraph, Dictionary<int, string>> paraToInsertMarkdown = new Dictionary<WParagraph, Dictionary<int, string>>();
        static void Main(string[] args)
        {          
            using (FileStream inputFileStream = new FileStream(Path.GetFullPath(@"Data/Template.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                //Opens an existing document from stream through constructor of `WordDocument` class.
                using (WordDocument document = new WordDocument(inputFileStream, FormatType.Automatic))
                {
                    //Creates mail merge events handler to replace merge field with HTML.
                    document.MailMerge.MergeField += new MergeFieldEventHandler(MergeFieldEvent);
                    //Gets data to perform mail merge.
                    DataTable table = GetDataTable();
                    //Performs the mail merge.
                    document.MailMerge.Execute(table);
                    //Append Markdown to paragraph.
                    InsertMarkdown();
                    //Removes mail merge events handler.
                    document.MailMerge.MergeField -= new MergeFieldEventHandler(MergeFieldEvent);
                    //Creates file stream.
                    using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"Output/Result.docx"), FileMode.Create, FileAccess.ReadWrite))
                    {
                        //Saves the Word document to file stream.
                        document.Save(outputFileStream, FormatType.Docx);
                    }
                    //Closes the document.
                    document.Close();
                }
            }
        }

        #region Helper methods
        /// <summary>
        /// Replaces merge field with Markdown string by using MergeFieldEventHandler.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="args"></param>
        public static void MergeFieldEvent(object sender, MergeFieldEventArgs args)
        {
            if (args.TableName.Equals("Markdown"))
            {
                if (args.FieldName.Equals("ProductList"))
                {
                    //Gets the current merge field owner paragraph.
                    WParagraph paragraph = args.CurrentMergeField.OwnerParagraph;
                    //Gets the current merge field index in the current paragraph.
                    int mergeFieldIndex = paragraph.ChildEntities.IndexOf(args.CurrentMergeField);
                    //Maintain Markdown in collection.
                    Dictionary<int, string> fieldValues = new Dictionary<int, string>();
                    fieldValues.Add(mergeFieldIndex, args.FieldValue.ToString());
                    //Maintain paragraph in collection.
                    paraToInsertMarkdown.Add(paragraph, fieldValues);
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
            DataTable dataTable = new DataTable("Markdown");
            dataTable.Columns.Add("CustomerName");
            dataTable.Columns.Add("Address");
            dataTable.Columns.Add("Phone");
            dataTable.Columns.Add("ProductList");
            DataRow datarow = dataTable.NewRow();
            dataTable.Rows.Add(datarow);
            datarow["CustomerName"] = "Nancy Davolio";
            datarow["Address"] = "59 rue de I'Abbaye, Reims 51100, France";
            datarow["Phone"] = "1-888-936-8638";
            //Markdown content.
            string markdown = "# Hello Markdown!\nThis is some **bold** text.";
            datarow["ProductList"] = markdown;
            return dataTable;
        }
        /// <summary>
        /// Append Markdown to paragraph.
        /// </summary>
        private static void InsertMarkdown()
        {
            //Iterates through each item in the dictionary.
            foreach (KeyValuePair<WParagraph, Dictionary<int, string>> dictionaryItems in paraToInsertMarkdown)
            {
                WParagraph paragraph = dictionaryItems.Key as WParagraph;
                Dictionary<int, string> values = dictionaryItems.Value as Dictionary<int, string>;
                //Iterates through each value in the dictionary.
                foreach (KeyValuePair<int, string> valuePair in values)
                {
                    int index = valuePair.Key;
                    string fieldValue = valuePair.Value;
                    // Convert the markdown string to bytes using UTF-8 encoding
                    byte[] contentBytes = Encoding.UTF8.GetBytes(fieldValue);

                    // Create a MemoryStream from the content bytes
                    using (MemoryStream memoryStream = new MemoryStream(contentBytes))
                    {
                        //Open the markdown Word document
                        using (WordDocument markdownDoc = new WordDocument(memoryStream, FormatType.Markdown))
                        {
                            TextBodyPart bodyPart = new TextBodyPart(paragraph.OwnerTextBody.Document);
                            BodyItemCollection m_bodyItems = bodyPart.BodyItems;
                            //Copy and paste the markdown at the same position of mergefield in Word document.
                            foreach (Entity entity in markdownDoc.LastSection.Body.ChildEntities)
                            {
                                m_bodyItems.Add(entity.Clone());
                            }
                            bodyPart.PasteAt(paragraph.OwnerTextBody, paragraph.OwnerTextBody.ChildEntities.IndexOf(paragraph), index);
                        }
                    }
                }
            }
            paraToInsertMarkdown.Clear();
        }
        #endregion
    }
}
