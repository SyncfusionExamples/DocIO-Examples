using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;

class Program
{
    static void Main(string[] args)
    {
        // Create an input file stream to open the document
        using (FileStream inputStream = new FileStream(Path.GetFullPath(@"Data/Input.docx"), FileMode.Open, FileAccess.Read))
        {
            // Create a new Word document instance from the input stream
            using (WordDocument document = new WordDocument(inputStream, FormatType.Docx))
            {
                // Change the font of the document text to "Times New Roman"
                ChangeFontName(document, "Times New Roman");

                // Create an output file stream to save the modified document
                using (FileStream outputStream = new FileStream(Path.GetFullPath(@"Output/Result.docx"), FileMode.OpenOrCreate, FileAccess.ReadWrite))
                {
                    // Save the modified document to the output stream in DOCX format
                    document.Save(outputStream, FormatType.Docx);
                }
            }
        }
    }

    /// <summary>
    /// Changes the font name of all text and paragraph marks in the Word document.
    /// </summary>
    /// <param name="document">The WordDocument object to modify.</param>
    /// <param name="fontName">The new font name to apply.</param>
    private static void ChangeFontName(WordDocument document, string fontName)
    {
        // Find all paragraphs by EntityType in the Word document.
        List<Entity> paragraphs = document.FindAllItemsByProperty(EntityType.Paragraph, null, null);

        // Change the font name for all paragraph marks (non-printing characters).
        for (int i = 0; i < paragraphs.Count; i++)
        {
            WParagraph paragraph = paragraphs[i] as WParagraph;
            paragraph.BreakCharacterFormat.FontName = fontName;
        }

        // Find all text ranges by EntityType in the Word document.
        List<Entity> textRanges = document.FindAllItemsByProperty(EntityType.TextRange, null, null);

        // Change the font name for all text content in the document.
        for (int i = 0; i < textRanges.Count; i++)
        {
            WTextRange textRange = textRanges[i] as WTextRange;
            textRange.CharacterFormat.FontName = fontName;
        }
    }
}
