using Syncfusion.DocIO.DLS;
using Syncfusion.DocIO;

namespace Add_Editable_range_to_specific_paragraph
{
    class Program
    {
        static void Main(string[] args)
        {
            // Load the Word document
            using (WordDocument document = new WordDocument(Path.GetFullPath("Data/Template.docx")))
            {
                // Access the first section of the document
                WSection section = document.Sections[0];
                // Insert an editable range at index 1 in the 2nd paragraph
                AddEditableRange(document, section.Paragraphs[1], 1);
                // Insert an editable range at index 2 in the 4th paragraph
                AddEditableRange(document, section.Paragraphs[3], 1);
                // Insert an editable range at index 4 in the 12th paragraph
                AddEditableRange(document, section.Paragraphs[11], 4);
                // Save the modified document to the new file
                document.Save(Path.GetFullPath(@"../../../Output/Result.docx"));
            }
        }
        /// <summary>
        /// Inserts an editable range at a specific index within a paragraph.
        /// </summary>
        /// <param name="document">The Word document instance</param>
        /// <param name="paragraph">The paragraph where the editable range will be inserted.</param>
        /// <param name="index">The index within the paragraph's child entities to insert the editable range.</param>
        private static void AddEditableRange(WordDocument document, WParagraph paragraph, int index)
        {
            // Create the start of the editable range
            EditableRangeStart editableRangeStart = new EditableRangeStart(document);
            // Insert the editable range start at the specified index in the paragraph
            paragraph.ChildEntities.Insert(index, editableRangeStart);
            // Set the editor group to allow everyone to edit this range
            editableRangeStart.EditorGroup = EditorType.Everyone;
            // Create the end of the editable range
            EditableRangeEnd editableRangeEnd = new EditableRangeEnd(document);
            // Insert the editable range end after the editable content 
            paragraph.ChildEntities.Insert(index + 2, editableRangeEnd);
        }
    }
}
