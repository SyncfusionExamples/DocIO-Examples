using Syncfusion.DocIO.DLS;
using Syncfusion.DocIO;
using System.Text.RegularExpressions;

using (FileStream inputFileStream = new FileStream(Path.GetFullPath(@"Data/Template.docx"), FileMode.Open, FileAccess.Read))
{
    // Create a WordDocument instance by loading the DOCX file from the file stream.
    using (WordDocument document = new WordDocument(inputFileStream, FormatType.Docx))
    {
        // Apply bold formatting to specific text using a regular expression.
        ApplyBoldFormatting(document);

        // Save the modified document to an output file.
        using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"Output/Result.docx"), FileMode.OpenOrCreate, FileAccess.ReadWrite))
        {
            // Save the modified Word document to the specified file path.
            document.Save(outputFileStream, FormatType.Docx);
        }
    }
}

/// <summary>
/// Applies bold formatting to text enclosed in <b>...</b> tags in the Word document.
/// </summary>
static void ApplyBoldFormatting(WordDocument document)
{
    // Define a regular expression to find all occurrences of <b>...</b>.
    Regex regex = new Regex("<b>(.*?)</b>");

    // Find all matches of the regex pattern in the document.
    TextSelection[] matches = document.FindAll(regex);

    // Iterate through each match found in the document.
    foreach (TextSelection match in matches)
    {
        // Get the entire text range of the matched content.
        WTextRange textRange = match.GetAsOneRange();

        // Extract the full text of the match.
        string fullText = textRange.Text;

        // Define the length of the opening tag <b>.
        int startTagLength = "<b>".Length;

        // Find the index of the closing tag </b>.
        int endTagIndex = fullText.LastIndexOf("</b>");

        // Extract the opening tag, bold text, and closing tag as separate strings.
        string startTag = fullText.Substring(0, startTagLength);
        string boldText = fullText.Substring(startTagLength, endTagIndex - startTagLength);
        string endTag = fullText.Substring(endTagIndex);

        // Create new text ranges for each part (opening tag, bold text, and closing tag).
        WTextRange startTextRange = CreateTextRange(textRange, startTag);
        WTextRange boldTextRange = CreateTextRange(textRange, boldText);
        WTextRange endTextRange = CreateTextRange(textRange, endTag);

        // Apply bold formatting to the text range containing the bold text.
        boldTextRange.CharacterFormat.Bold = true;

        // Replace the original text range with the newly created text ranges in the paragraph.
        WParagraph paragraph = textRange.OwnerParagraph;

        // Get the index of the original text range within the paragraph.
        int index = paragraph.ChildEntities.IndexOf(textRange);

        // Remove the original text range from the paragraph.
        paragraph.ChildEntities.RemoveAt(index);

        // Insert the new text ranges (in order: closing tag, bold text, opening tag) into the paragraph.
        paragraph.ChildEntities.Insert(index, endTextRange);
        paragraph.ChildEntities.Insert(index, boldTextRange);
        paragraph.ChildEntities.Insert(index, startTextRange);
    }
}

/// <summary>
/// Creates a new text range with the specified text, copying formatting from the original range.
/// </summary>
static WTextRange CreateTextRange(WTextRange original, string text)
{
    // Clone the original text range to preserve its formatting.
    WTextRange newTextRange = original.Clone() as WTextRange;

    // Set the text for the new text range.
    newTextRange.Text = text;

    return newTextRange;
}
