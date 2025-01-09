using Syncfusion.DocIO.DLS;
using Syncfusion.DocIO;

namespace Change_format_after_append_html
{
    internal class Program
    {
        static void Main(string[] args)
        {
            // Open the Word document.
            using (FileStream inputStream = new FileStream(Path.GetFullPath(@"Data/Template.docx"), FileMode.Open, FileAccess.Read))
            {
                using (WordDocument document = new WordDocument(inputStream, FormatType.Docx))
                {
                    // Access the first section of the document.
                    WSection section = document.Sections[0];

                    // Get the index of the last paragraph.
                    int paraIndex = section.Body.Paragraphs.Count - 1;

                    // Append HTML content to the document with formatting in <p> tag.
                    document.LastParagraph.AppendHTML("<p style='color:blue; font-weight:bold;'>The Giant</p><p style='color:green; font-style:italic;'>Panda</p>");

                    // Iterate through the paragraphs in the section's body.
                    for (int i = paraIndex; i < section.Body.ChildEntities.Count; i++)
                    {
                        // Get the paragraph and check if it's not null.
                        WParagraph paragraph = section.Body.ChildEntities[i] as WParagraph;
                        if (paragraph != null)
                        {
                            // Set the paragraph formatting spacing to 0.
                            paragraph.ParagraphFormat.BeforeSpacing = 0;
                            paragraph.ParagraphFormat.AfterSpacing = 0;

                            // Iterate through the items in the paragraph to chnage formatting.
                            foreach (var item in paragraph.ChildEntities)
                            {
                                if (item is WTextRange textRange)
                                    textRange.CharacterFormat.TextColor = Syncfusion.Drawing.Color.DarkGreen; // Change text color
                            }
                        }
                    }
                    // Save the modified document.
                    using (FileStream outputStream = new FileStream(Path.GetFullPath(@"Output/Result.docx"), FileMode.Create, FileAccess.Write))
                    {
                        document.Save(outputStream, FormatType.Docx);
                    }
                }
            }
        }
    }
}
