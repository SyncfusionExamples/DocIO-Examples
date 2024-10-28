// Load the existing Word document
using Syncfusion.DocIO.DLS;
using Syncfusion.DocIO;
using Syncfusion.Drawing;

using (FileStream inputStream = new FileStream(Path.GetFullPath("Data/Input.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
{
    // Open the input Word document
    WordDocument document = new WordDocument(inputStream, FormatType.Docx);
    // Iterate through all the styles in the document
    foreach (Style style in document.Styles)
    {
        // Check if the style is a Paragraph style
        if (style.StyleType == StyleType.ParagraphStyle)
        {
            // Cast the style to WParagraphStyle and modify the text color
            WParagraphStyle paraStyle = style as WParagraphStyle;
            if (paraStyle != null)
            {
                paraStyle.CharacterFormat.TextColor = Color.Purple;
            }
        }
        // Check if the style is a Character style
        else if (style.StyleType == StyleType.CharacterStyle)
        {
            // Cast the style to WCharacterStyle and modify the text color
            WCharacterStyle charStyle = style as WCharacterStyle;
            if (charStyle != null)
            {
                charStyle.CharacterFormat.TextColor = Color.Red;
            }
        }
    }
    // Save the modified document
    using (FileStream outputStream = new FileStream(Path.GetFullPath("Output/Result.docx"), FileMode.Create, FileAccess.Write))
    {
        document.Save(outputStream, FormatType.Docx);
    }
}