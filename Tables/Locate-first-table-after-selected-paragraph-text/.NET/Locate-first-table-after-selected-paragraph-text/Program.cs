using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
namespace Locate_first_table_after_selected_paragraph_text
{
    class Program
    {
        static void Main(string[] args)
        {
            using (WordDocument document = new WordDocument(Path.GetFullPath(@"../../../Data/Template.docx")))
            {
                //Finds the first occurrence of a particular text in the document
                TextSelection textSelection = document.Find("About the Suppliers", false, true);
                //Gets the found text as single text range
                WTextRange textRange = textSelection.GetAsOneRange();
                //Gets the owner paragraph of the selected text
                WParagraph paragraph = textRange.OwnerParagraph;
                //Gets the owner textbody of the selected paragraph
                WTextBody textbody = paragraph.OwnerTextBody;
                //Gets the index of the selected paragraph
                int paraIndex = textbody.ChildEntities.IndexOf(paragraph);

                //Gets the first table after the selected paragraph text
                WTable firstTableAfterSelection = GetFirstTableAfterSelection(textbody, paraIndex);
                if (firstTableAfterSelection != null)
                {
                    //To do anything;
                }
                //Save the Word document
                document.Save(Path.GetFullPath(@"../../../Output/output.docx"), FormatType.Docx);
            }
        }
        /// <summary>
        /// Returns the first table after the selected paragraph text.
        /// </summary>
        /// <param name="body">The text body.</param>
        /// <param name="index">The index of the selected paragraph.</param>
        /// <returns></returns>
        static WTable GetFirstTableAfterSelection(WTextBody body, int paraIndex)
        {
            //Iterates through the body items to get the first table after the selected paragraph text.
            for (int i = paraIndex + 1; i < body.ChildEntities.Count; i++)
            {
                if (body.ChildEntities[i] is WTable)
                {
                    return body.ChildEntities[i] as WTable;
                }
            }
            //Returns null, if there is no table after the selected paragraph text in the Word document.
            return null;
        }
    }
}