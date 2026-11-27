using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;

namespace Toggle_field_codes
{
    class Program
    {
        static void Main(string[] args)
        {
            using (FileStream inputStream = new FileStream(Path.GetFullPath("../../../Data/Template.docx"), FileMode.Open, FileAccess.Read))
            {
                //Open word document instance
                using (WordDocument document = new WordDocument(inputStream, FormatType.Docx))
                {
                    //Access field that need to be toggled.
                    WField field = document.FindAllItemsByProperty(EntityType.Field, null, null)[0] as WField;
                    if (field != null)
                        ToggleFieldCodes(document, field);
                    //Save modified document
                    using (FileStream outputStream = new FileStream(Path.GetFullPath("../../../Output/output.docx"), FileMode.Create, FileAccess.ReadWrite))
                    {
                        document.Save(outputStream, FormatType.Docx);
                    }
                }
            }
        }
        #region ToggleFieldCode
        /// <summary>
        /// Converts the specified field result into its field code representation by
        /// extracting the field code content and replacing the field result.
        /// </summary>
        private static void ToggleFieldCodes(WordDocument document, WField field)
        {
            // Get the paragraph and section where the field exists.
            WParagraph ownerPara = field.OwnerParagraph;
            WSection ownerSection = (WSection)ownerPara.OwnerTextBody.Owner;

            // Get the index of the section and paragraph.
            int sectionIndex = document.Sections.IndexOf(ownerSection);
            int bodyIndex = ownerSection.Paragraphs.IndexOf(ownerPara);
            int fieldIndex = ownerPara.ChildEntities.IndexOf(field);

            // Stack to keep track of nested fields.
            bool isAppendText = true;
            Stack<Entity> fieldStack = new Stack<Entity>();

            // Create a new document part for replacing the bookmark content.
            WordDocumentPart documentPart = new WordDocumentPart();
            bool isSeparatorReached = false;

            // Process each section until all fields are handled.
            for (int i = sectionIndex; i < document.Sections.Count; i++)
            {
                WSection currentSection = document.Sections[i];
                WSection section = currentSection.Clone();
                section.Body.ChildEntities.Clear();
                documentPart.Sections.Add(section);
                //Process section and its child entities till field end is found.
                ProcessBody(document, bodyIndex, fieldIndex, currentSection, documentPart, section, ref fieldStack, ref isAppendText, ref isSeparatorReached);
                if (fieldStack.Count == 0)
                    return;

                // Reset indices for further processing.
                bodyIndex = 0;
                fieldIndex = 0;
            }
        }
        
        /// <summary>
        /// Processes the section content, identifies field boundaries, extracts field
        /// code content, and builds the document part used for field code replacement.
        /// </summary>
        private static void ProcessBody(WordDocument document, int bodyIndex, int fieldIndex, WSection currentSection, WordDocumentPart documentPart, WSection sectionToBeInserted, ref Stack<Entity> fieldStack, ref bool isAppendText, ref bool isSeparatorReached)
        {
            for (int j = bodyIndex; j < currentSection.Body.ChildEntities.Count; j++)
            {
                Entity entity = currentSection.Body.ChildEntities[j];
                if (entity is WParagraph paragraph)
                {
                    WParagraph paraToBeInserted = (WParagraph)paragraph.Clone();
                    paraToBeInserted.ChildEntities.Clear();
                    // New para need to added only when isAppendText is true
                    if (isAppendText)
                        sectionToBeInserted.Body.ChildEntities.Add(paraToBeInserted);
                    //Iterate the paragraph child entities
                    for (int i = fieldIndex; i < paragraph.ChildEntities.Count; i++)
                    {
                        Entity paraItem = paragraph.ChildEntities[i];
                        //When it encounters field, push in stack and append }
                        if (paraItem is WField field)
                        {
                            fieldStack.Push(field);
                            //Append "{" when isAppendText should be true and parent field separator should not be reached
                            if (!isSeparatorReached && isAppendText)
                            {
                                paraToBeInserted.AppendText("{");
                                isAppendText = true;
                            }
                        }
                        else if (paraItem is WFieldMark fieldMark)
                        {
                            if (fieldMark.Type == FieldMarkType.FieldSeparator)
                            {
                                if (isAppendText)
                                    paraToBeInserted.AppendText("}");
                                //Entities should not append in document part after field separator
                                isAppendText = false;
                                //When it is parent field separator, add bookmark start next to it
                                if (fieldStack.Count == 1)
                                {
                                    int separatorIndex = paragraph.ChildEntities.IndexOf(fieldMark);
                                    BookmarkStart bkmkStart = new BookmarkStart(document, "bkmk");
                                    paragraph.ChildEntities.Insert(separatorIndex + 1, bkmkStart);
                                    isSeparatorReached = true;
                                }
                            }
                            else
                            {
                                //When it is parent field end, add bookmark end before to it
                                if (fieldStack.Count == 1)
                                {
                                    int endIndex = paragraph.ChildEntities.IndexOf(fieldMark);
                                    BookmarkEnd bkmkEnd = new BookmarkEnd(document, "bkmk");
                                    paragraph.ChildEntities.Insert(endIndex, bkmkEnd);
                                    fieldStack.Pop();
                                    break;
                                }
                                fieldStack.Pop();
                                isAppendText = true;
                            }
                        }
                        else if (paraItem is WTextRange textRange)
                        {
                            if (isAppendText)
                                paraToBeInserted.ChildEntities.Add(textRange.Clone());
                        }
                    }
                }
                else if (entity is WTable table)
                {
                    // Clone the table and add it to the section.
                    if (isAppendText)
                    {
                        WTable tableToBeInserted = table.Clone();
                        sectionToBeInserted.Body.ChildEntities.Add(tableToBeInserted);
                    }
                }
                if (fieldStack.Count == 0)
                {
                    //When fields are processed, replace bookmark content
                    ReplaceBookmarkContent(document, "bkmk", documentPart);
                    return;
                }
                fieldIndex = 0;
            }
        }
        
        /// <summary>
        /// Replaces the specified bookmark content with the generated document part
        /// and removes the bookmark after replacement.
        /// </summary>
        private static void ReplaceBookmarkContent(WordDocument document, string bookmarkName, WordDocumentPart documentPart)
        {
            //Creates the bookmark navigator instance to access the bookmark.
            BookmarksNavigator bookmarkNavigator = new BookmarksNavigator(document);
            //Moves the virtual cursor to the location before the end of the bookmark "bkmk".
            bookmarkNavigator.MoveToBookmark(bookmarkName);
            //Replaces the bookmark content with word body part.
            bookmarkNavigator.ReplaceContent(documentPart);
            //Close the WordDocumentPart instance.
            documentPart.Close();
            //Remove bookmark from the document after replacement
            document.Bookmarks.Remove(bookmarkNavigator.CurrentBookmark);
        }
        #endregion

        #region UnToggle Field Codes
        /// <summary>
        /// Restores the field result by updating the field or converting merge field
        /// codes back to their display format.
        /// </summary>
        private static void UnToggleFieldCodes(WField field)
        {
            if (field is WMergeField)
            {
                WMergeField mergeField = field as WMergeField;
                string fieldResult = "«" + mergeField.FieldName + "»";
                mergeField.Text = fieldResult;
            }
            else
                field.Update();
        }
        #endregion
    }
}
