using Syncfusion.DocIO.DLS;
using Syncfusion.DocIO;
using System.Text.RegularExpressions;

namespace Find_and_apply_bold_to_replaced_content
{
    internal class Program
    {
        static void Main(string[] args)
        {
            // Open the input Word document from the specified path.
            using (FileStream docStream = new FileStream(Path.GetFullPath(@"Data/Input.docx"), FileMode.Open, FileAccess.Read))
            {
                // Load the Word document into the WordDocument object.
                using (WordDocument document = new WordDocument(docStream, FormatType.Docx))
                {
                    // Find the text pattern using a regular expression.
                    TextSelection selection = document.Find(new Regex("^<<([^>]*)>>"));
                    // Get the selected text as a single text range.
                    WTextRange textRange = selection.GetAsOneRange();
                    // Get the text body and paragraph of selected text.
                    WTextBody body = textRange.OwnerParagraph.OwnerTextBody;
                    WParagraph paragraph = textRange.OwnerParagraph;
                    // Get the index of the paragraph within the text body.
                    int paraIndex = body.ChildEntities.IndexOf(paragraph);
                    // Replace the selected text with the replace content.
                    document.Replace(selection.SelectedText, "Adventure Works Cycles", true, true);
                    // Search for the word within the current paragraph.
                    TextSelection replacedSelection = paragraph.Find("Works", true, false);
                    if (replacedSelection != null)
                    {
                        // Get the selected text as a single text range.
                        WTextRange replacedTextRange = replacedSelection.GetAsOneRange();
                        // Apply bold formatting to the text range.
                        replacedTextRange.CharacterFormat.Bold = true;
                    }
                    // Save the modified document to the specified output path.
                    using (FileStream outputStream = new FileStream(Path.GetFullPath(@"Output/Result.docx"), FileMode.Create, FileAccess.Write))
                    {
                        document.Save(outputStream, FormatType.Docx);
                    }
                }
            }
        }
    }
}