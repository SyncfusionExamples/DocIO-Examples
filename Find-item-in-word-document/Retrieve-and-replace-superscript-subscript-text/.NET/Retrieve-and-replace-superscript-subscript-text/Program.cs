using Syncfusion.DocIO.DLS;

namespace Retrieve_and_replace_superscript_subscript_text
{
    class Program
    {
        public static void Main(string[] args)
        {
            // Load the existing Word document.
            WordDocument document = new WordDocument(Path.GetFullPath(@"Data\Input.docx"));
            // Replace all superscript text and maintain as superscript
            ReplaceSuperscriptAndSubscriptText(document, "SuperScript", false);
            // Replace all subscript text but convert it into normal text.
            ReplaceSuperscriptAndSubscriptText(document, "SubScript", true);
            // Save the word document
            document.Save(Path.GetFullPath("../../../Output/Output.docx"));
            // Close the document.
            document.Close();
        }
        /// <summary>
        /// Replaces superscript or subscript text in a Word document.
        /// </summary>
        /// <param name="document">The Word document to process.</param>        
        /// <param name="subSuperScriptType">Type of script</param>
        /// <param name="displayNormalText">True if the replaced text should be converted to normal text; false to keep formatting.</param>
        static void ReplaceSuperscriptAndSubscriptText(WordDocument document, string subSuperScriptType, bool displayNormalText)
        {
            // Find all text ranges with the given superscript or subscript formatting.
            List<Entity> textRangesWithsubsuperScript = document.FindAllItemsByProperty(EntityType.TextRange, "CharacterFormat.SubSuperScript", subSuperScriptType);
            for (int i = 0; i < textRangesWithsubsuperScript.Count; i++)
            {
                // Cast the entity to a text range.
                WTextRange textRange = textRangesWithsubsuperScript[i] as WTextRange;
                // Replace the existing text with new content
                textRange.Text = $"<{subSuperScriptType}> {textRange.Text} </{subSuperScriptType}>";
                // If the replaced content displayed as normal text
                if(displayNormalText)
                {
                    // Set SubSuperScript as none.
                    textRange.CharacterFormat.SubSuperScript = SubSuperScript.None;
                }
            }
        }
    }
}