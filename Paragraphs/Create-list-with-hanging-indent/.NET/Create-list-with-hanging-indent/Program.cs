using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;

namespace Create_list_with_hanging_indent
{
    internal class Program
    {
        static void Main(string[] args)
        {
            // Create a new Word document.
            using (WordDocument document = new WordDocument())
            {
                // Add a section to the document.
                IWSection section = document.AddSection();

                // Add a bulleted list item with a hanging indent.
                IWParagraph bulletParagraph1 = section.AddParagraph();
                bulletParagraph1.ListFormat.ApplyDefBulletStyle();
                bulletParagraph1.AppendText("First").CharacterFormat.FontSize = 12;
                bulletParagraph1.ParagraphFormat.FirstLineIndent = -18; // Hanging indent for bullet alignment.

                // Add a normal paragraph with a first-line indent.
                IWParagraph normalParagraph = section.AddParagraph();
                normalParagraph.AppendText("Sample text with first-line indent.").CharacterFormat.FontSize = 12;
                normalParagraph.ParagraphFormat.FirstLineIndent = 35;

                // Add another bulleted list item with a hanging indent.
                IWParagraph bulletParagraph2 = section.AddParagraph();
                bulletParagraph2.ListFormat.ApplyDefBulletStyle();
                bulletParagraph2.AppendText("Second").CharacterFormat.FontSize = 12;
                bulletParagraph2.ParagraphFormat.FirstLineIndent = -18; // Hanging indent for bullet alignment.

                // Save the document.
                using (FileStream outputFileStream = new FileStream(Path.GetFullPath("Output/Result.docx"), FileMode.Create, FileAccess.Write))
                {
                    document.Save(outputFileStream, FormatType.Docx);
                }
            }
        }
    }
}
