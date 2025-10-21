using Syncfusion.DocIO.DLS;

namespace Find_Text_and_Make_it_Editable
{
    public class Program
    {
        public static void Main(string[] args)
        {
            //Loads template document
            WordDocument document = new WordDocument(Path.GetFullPath(@"Data/Template.docx"));
            // Find all instances of the target text (text, case-insensitive, whole word match)
            TextSelection[] textSelections = document.FindAll("Adventure Works Cycles", false, true);

            // Append editable ranges around each matched text.
            foreach (var selection in textSelections)
            {
                AppendEditableRange(selection);
            }
            // save and close the document
            document.Save(Path.GetFullPath(Path.GetFullPath(@"../../../Output/Result.docx")));
            document.Close();
        }

        public static void AppendEditableRange(TextSelection selection)
        {
            // Get the text range from the selection.
            WTextRange[] textRanges = selection.GetRanges();
            if (textRanges.Length > 0)
            {
                // Get the first and last text ranges in the selection.
                WTextRange startTextRange = textRanges[0];
                WTextRange endTextRange = textRanges[textRanges.Length - 1];
                // Get the paragraph that owns the start text range.
                WParagraph paragraph = startTextRange.OwnerParagraph;
                // Create a new EditableRangeStart and assign it to everyone by default.
                EditableRangeStart editableRangeStart = new EditableRangeStart(paragraph.Document);
                editableRangeStart.EditorGroup = EditorType.Everyone;
                // Find the index of the start text range within the paragraph's child entities.
                int startTextRangeIndex = paragraph.ChildEntities.IndexOf(startTextRange);
                // Insert the editable range start before the start text range.
                paragraph.ChildEntities.Insert(startTextRangeIndex, editableRangeStart);
                // Create a new EditableRangeEnd linked to the start.
                EditableRangeEnd editableRangeEnd = new EditableRangeEnd(paragraph.Document, editableRangeStart);
                // Find the index of the end text range and insert the editable range end after it.
                int endTextRangeIndex = paragraph.ChildEntities.IndexOf(endTextRange);
                paragraph.ChildEntities.Insert(endTextRangeIndex + 1, editableRangeEnd);
            }
        }
    }
}
