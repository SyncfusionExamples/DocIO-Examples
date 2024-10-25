using Syncfusion.DocIO.DLS;
using Syncfusion.DocIO;
using System.Text.RegularExpressions;

// Open the DOCX file from the file stream.
using (FileStream fileStream1 = new FileStream(Path.GetFullPath(@"Data/Template.docx"), FileMode.Open, FileAccess.Read))
{
    // Create a WordDocument instance by loading the DOCX file from the file stream.
    using (WordDocument document = new WordDocument(fileStream1, FormatType.Docx))
    {
        // Apply bold formatting to specific text using a regular expression.
        ApplyBoldUsingRegex(document);

        // Save the modified document to an output file.
        using (FileStream outputStream1 = new FileStream(Path.GetFullPath(@"Output/Result.docx"), FileMode.OpenOrCreate, FileAccess.ReadWrite))
        {
            document.Save(outputStream1, FormatType.Docx);
        }
    }
}

// Method to apply bold formatting to text matching a specific regex pattern.
static void ApplyBoldUsingRegex(WordDocument document)
{
    int startTagIndex = 0;  // To store the index of the start tag.
    int endTagIndex = 0;    // To store the index of the end tag.
    WParagraph para = null; // Reference to the paragraph containing the tags.
    WTextRange startTextRange = null; // Text range for the start tag.
    WTextRange endTextRange = null;   // Text range for the end tag.

    // Find the text that matches the <b>...</b> pattern using a regex.
    TextSelection textSelection = document.Find(new Regex("<b>(.*)</b>"));

    // Get all text ranges that match the regex pattern.
    WTextRange[] textRanges = textSelection.GetRanges();

    // Iterate through each matched text range.
    for (int i = 0; i < textRanges.Length; i++)
    {
        WTextRange textRange = textRanges[i];

        // If the text range contains both <b> and </b> tags.
        if (i == 0 || i == textRanges.Length - 1)
        {
            if (textRange.Text.Contains("<b>") && textRange.Text.Contains("</b>"))
            {
                // Process the text with both start and end tags.
                ProcessTextWithStartAndEndTags(textRange);
            }
            else if (textRange.Text.Contains("<b>") || textRange.Text.Contains("</b>"))
            {
                // Process the text with only the start or end tag.
                ProcessTextWithPartialTags(textRange, ref para, ref startTextRange, ref endTextRange, ref startTagIndex, ref endTagIndex);
            }
        }
        else
        {
            // Apply bold formatting to the text between <b> and </b>.
            textRange.CharacterFormat.Bold = true;
        }
    }

    // Insert the start and end text ranges if applicable.
    if (para != null)
    {
        para.ChildEntities.Insert(startTagIndex, startTextRange);
        para.ChildEntities.Insert(endTagIndex + 1, endTextRange);
    }
}

// Process text that contains both <b> and </b> tags.
static void ProcessTextWithStartAndEndTags(WTextRange textRange)
{
    // Find the indexes of the start and end tags.
    int startIndex = textRange.Text.IndexOf("<b>") + 3;
    int endIndex = textRange.Text.IndexOf("</b>");

    // Create text ranges for the text before the start tag and after the end tag.
    WTextRange startTextRange1 = CreateTextRange(textRange, textRange.Text.Substring(0, startIndex));
    WTextRange endTextRange1 = CreateTextRange(textRange, textRange.Text.Substring(endIndex));

    // Extract and format the text within the tags.
    string boldText = textRange.Text.Substring(startIndex, endIndex - startIndex);
    textRange.Text = boldText;
    textRange.CharacterFormat.Bold = true;

    // Insert the start and end text ranges into the paragraph.
    WParagraph para = textRange.OwnerParagraph;
    int index = para.ChildEntities.IndexOf(textRange);
    para.ChildEntities.Insert(index, startTextRange1);
    para.ChildEntities.Insert(index + 2, endTextRange1);
}

// Process text that contains only <b> or only </b> tags.
static void ProcessTextWithPartialTags(WTextRange textRange, ref WParagraph para, ref WTextRange startTextRange, ref WTextRange endTextRange, ref int startTagIndex, ref int endTagIndex)
{
    if (textRange.Text.Contains("<b>"))
    {
        // If the text contains only the start tag <b>.
        int startIndex = textRange.Text.IndexOf("<b>") + 3;
        startTextRange = CreateTextRange(textRange, textRange.Text.Substring(0, startIndex));
        para = textRange.OwnerParagraph;
        startTagIndex = para.ChildEntities.IndexOf(textRange);
        textRange.Text = textRange.Text.Replace("<b>", "");
    }
    else if (textRange.Text.Contains("</b>"))
    {
        // If the text contains only the end tag </b>.
        int endIndex = textRange.Text.IndexOf("</b>");
        endTextRange = CreateTextRange(textRange, textRange.Text.Substring(endIndex));
        para = textRange.OwnerParagraph;
        endTagIndex = para.ChildEntities.IndexOf(textRange);
        textRange.Text = textRange.Text.Replace("</b>", "");
    }

    // Apply bold formatting to the text.
    textRange.CharacterFormat.Bold = true;
}

// Utility method to create a new WTextRange with the given text.
static WTextRange CreateTextRange(WTextRange original, string text)
{
    // Clone the original text range and update its text.
    WTextRange newTextRange = original.Clone() as WTextRange;
    newTextRange.Text = text;
    return newTextRange;
}