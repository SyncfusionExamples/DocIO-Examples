using Syncfusion.DocIO.DLS;
using Syncfusion.DocIO;
using System;
using System.IO;
using System.Collections.Generic;
using System.Data;

namespace Apply_merge_field_formatting_to_InsertHTML
{
    internal class Program
    {
        static void Main(string[] args)
        {
            using (FileStream fileStreamPath = new FileStream(Path.GetFullPath(@"../../../Data/Template.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                //Creates a new Word document.
                using (WordDocument document = new WordDocument(fileStreamPath, FormatType.Automatic))
                {
                    Dictionary<WParagraph, List<MergeFieldInfo>> paraToInsertHTML = new Dictionary<WParagraph, List<MergeFieldInfo>>();
                    //Creates the mail merge events handler to replace merge field with HTML.
                    MergeFieldEventHandler mergeHandler = (sender, eventArgs) => MergeFieldEvent(eventArgs, paraToInsertHTML);
                    document.MailMerge.MergeField += mergeHandler;
                    //Gets data to perform the mail merge.
                    DataTable table = GetDataTable();
                    //Performs the mail merge.
                    document.MailMerge.Execute(table);
                    // If all the paragraphs need to ignore the styles of the merge field means, please set false to mergeFieldStyle boolean.
                    bool mergeFieldStyle = true;
                    //Append HTML to paragraph.
                    InsertHtml(paraToInsertHTML, mergeFieldStyle);
                    //Removes the mail merge events handler.
                    document.MailMerge.MergeField -= mergeHandler;
                    document.Save(@"../../../Output/output.docx");
                    paraToInsertHTML.Clear();
                }
            }
        }

        /// <summary>
        /// Handles the mail merge field event and collects XHTML content along with
        /// the merge field position and character formatting information for later insertion.
        /// </summary>
        /// <param name="args">Contains information about the current merge field being processed.</param>
        /// <param name="paraToInsertHTML">
        /// Dictionary that maps paragraphs to their corresponding XHTML insertion details.
        /// </param>
        public static void MergeFieldEvent(MergeFieldEventArgs args, Dictionary<WParagraph, List<MergeFieldInfo>> paraToInsertHTML)
        {
            if (args.TableName.Equals("HTML"))
            {
                if (args.FieldName.Equals("DisplayHTML"))
                {
                    //Gets the current merge field owner paragraph.
                    WParagraph paragraph = args.CurrentMergeField.OwnerParagraph;
                    //Gets the current merge field index in the current paragraph.
                    int mergeFieldIndex = paragraph.ChildEntities.IndexOf(args.CurrentMergeField);

                    if (!paraToInsertHTML.TryGetValue(paragraph, out List<MergeFieldInfo> fieldList))
                    {
                        fieldList = new List<MergeFieldInfo>();
                        paraToInsertHTML.Add(paragraph, fieldList);
                    }

                    fieldList.Add(new MergeFieldInfo
                    {
                        Index = mergeFieldIndex,
                        Html = args.FieldValue.ToString(),
                        CharacterFormat = args.CurrentMergeField.CharacterFormat
                    });

                    //Set field value as empty.
                    args.Text = string.Empty;
                }
            }
        }

        /// <summary>
        /// Inserts XHTML content at the corresponding merge field locations and optionally
        /// applies the merge field character formatting to all newly inserted content,
        /// including content spanning multiple paragraphs.
        /// </summary>
        /// <param name="paraToInsertHTML">
        /// Collection containing paragraphs and XHTML insertion details.
        /// </param>
        /// <param name="mergeFieldStyle">
        /// True to apply the merge field character formatting to inserted XHTML content;
        /// otherwise, false.
        /// </param>
        public static void InsertHtml(Dictionary<WParagraph, List<MergeFieldInfo>> paraToInsertHTML, bool mergeFieldStyle)
        {
            //Iterates through each item in the dictionary.
            foreach (KeyValuePair<WParagraph, List<MergeFieldInfo>> dictionaryItems in paraToInsertHTML)
            {
                WParagraph paragraph = dictionaryItems.Key;
                List<MergeFieldInfo> fieldList = dictionaryItems.Value;

                //Iterates through each merge field in the paragraph.
                foreach (MergeFieldInfo fieldInfo in fieldList)
                {
                    //Get count of paragraph items.
                    int paraItemsCountBefore = paragraph.ChildEntities.Count;

                    //Get count of owner body items.
                    int ownerItemsCountBefore = paragraph.OwnerTextBody.ChildEntities.Count;

                    //Inserts an HTML string at the same position of mergefield in a Word document.
                    paragraph.OwnerTextBody.InsertXHTML(fieldInfo.Html, paragraph.OwnerTextBody.ChildEntities.IndexOf(paragraph), fieldInfo.Index);

                    //Get count of paragraph items.
                    int paraItemsCountAfter = paragraph.ChildEntities.Count;
                    //Get count of owner body items.
                    int ownerItemsCountAfter = paragraph.OwnerTextBody.ChildEntities.Count;

                    if (mergeFieldStyle)
                    {
                        //If difference in paragraph child items count means then apply merge field character format to newly added items.
                        if (paraItemsCountBefore != paraItemsCountAfter)
                        {
                            ApplyFormat(paragraph, paraItemsCountAfter - paraItemsCountBefore, paraItemsCountAfter, fieldInfo.Index, fieldInfo.CharacterFormat);
                        }
                        //Iterate the owner body items and apply the merge field character format to newly added paragraphs.
                        if (ownerItemsCountBefore != ownerItemsCountAfter)
                        {
                            int diffCount = ownerItemsCountAfter - ownerItemsCountBefore;
                            int j = 0;
                            for (int i = paragraph.OwnerTextBody.ChildEntities.IndexOf(paragraph); i < ownerItemsCountAfter; i++)
                            {
                                if (j <= diffCount && paragraph.OwnerTextBody.ChildEntities[i] is WParagraph insertedPara)
                                    ApplyFormat(insertedPara, insertedPara.ChildEntities.Count, insertedPara.ChildEntities.Count, 0, fieldInfo.CharacterFormat);
                                j++;
                            }
                        }
                    }
                }
            }
            paraToInsertHTML.Clear();
        }

        /// <summary>
        /// Applies the specified character formatting to the newly inserted text ranges
        /// within the given paragraph.
        /// </summary>
        /// <param name="paragraph">The paragraph containing the inserted content.</param>
        /// <param name="diffCount">
        /// Number of newly inserted paragraph items to format.
        /// </param>
        /// <param name="totalCount">
        /// Total number of child entities present in the paragraph.
        /// </param>
        /// <param name="index">
        /// Starting index from which formatting should be applied.
        /// </param>
        /// <param name="format">
        /// Character formatting to apply to the inserted text.
        /// </param>
        private static void ApplyFormat(WParagraph paragraph, int diffCount, int totalCount, int index, WCharacterFormat format)
        {
            int j = 0;
            //Iterate the paragraph items (newly inserted) from merge field index.
            for (int i = index; i < totalCount; i++)
            {
                //Apply merge field character format to newly added item.
                if (j < diffCount && paragraph.ChildEntities[i] is WTextRange textRange)
                {
                    textRange.ApplyCharacterFormat(format);
                }
                j++;
            }
        }

        /// <summary>
        /// Creates and returns a data source used for mail merge execution.
        /// The XHTML content is read from an external file and assigned to the merge field.
        /// </summary>
        /// <returns>
        /// A DataTable containing XHTML content for mail merge processing.
        /// </returns>
        private static DataTable GetDataTable()
        {
            DataTable dataTable = new DataTable("HTML");
            dataTable.Columns.Add("DisplayHTML");
            DataRow datarow = dataTable.NewRow();
            dataTable.Rows.Add(datarow);
            //Reads HTML string from the file.
            string htmlString = File.ReadAllText(@"../../../Data/HTML.txt");
            datarow["DisplayHTML"] = htmlString;
            return dataTable;
        }

        /// <summary>
        /// Stores information required to replace a merge field with XHTML content
        /// and preserve the original merge field character formatting.
        /// </summary>
        internal class MergeFieldInfo
        {
            /// <summary>
            /// Gets or sets the index position of the merge field within the paragraph.
            /// </summary>
            public int Index { get; set; }
            /// <summary>
            /// Gets or sets the XHTML content that will be inserted at the merge field location.
            /// </summary>
            public string Html { get; set; }
            /// <summary>
            /// Gets or sets the character formatting of the merge field that should be
            /// applied to the inserted XHTML content.
            /// </summary>
            public WCharacterFormat CharacterFormat { get; set; }
        }

    }
}
