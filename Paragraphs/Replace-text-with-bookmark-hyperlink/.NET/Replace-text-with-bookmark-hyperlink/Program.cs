using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;

namespace Replace_text_with_bookmark_hyperlink
{
    class Program
    {
        static void Main(string[] args)
        {
            using (FileStream fileStreamPath = new FileStream(Path.GetFullPath(@"Data/Template.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                //Open an existing Word document
                using (WordDocument document = new WordDocument(fileStreamPath, FormatType.Automatic))
                {
                    string textToReplace = "Mexico";
                    // Initialize text body part
                    TextBodyPart textBodyPart = new TextBodyPart(document);
                    // Initialize new paragraph
                    WParagraph paragraph = new WParagraph(document);
                    // Add the paragraph to the text body part
                    textBodyPart.BodyItems.Add(paragraph);
                    // Append a bookmark hyperlink to the paragraph
                    paragraph.AppendHyperlink("bookmark", "Mexico", HyperlinkType.Bookmark);
                    // Replace all occurrences of the target text with the text body part
                    document.Replace(textToReplace, textBodyPart, false, true);
                    // Clear the text body part
                    textBodyPart.Clear();
                    // Create the output file stream
                    using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"Output/Result.docx"), FileMode.Create, FileAccess.ReadWrite))
                    {
                        // Save the Word document to the output stream
                        document.Save(outputFileStream, FormatType.Docx);
                    }
                }
            }
        }
    }
}

