using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.Drawing;
using System.IO;

namespace Append_text_to_paragraph
{
    class Program
    {
        static void Main(string[] args)
        {
            //Creates a new Word document.
            using (WordDocument document = new WordDocument())
            {
                //Adds new section to the document.
                IWSection section = document.AddSection();
                //Adds new paragraph to the section.
                IWParagraph firstParagraph = section.AddParagraph();
                //Adds new text to the paragraph.
                IWTextRange firstText = firstParagraph.AppendText("A new text is added to the paragraph.");
                firstText.CharacterFormat.FontSize = 14;
                firstText.CharacterFormat.Bold = true;
                firstText.CharacterFormat.TextColor = Color.Green;
                //Creates file stream.
                using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"Output/Result.docx"), FileMode.Create, FileAccess.ReadWrite))
                {
                    //Saves the Word document to file stream.
                    document.Save(outputFileStream, FormatType.Docx);
                }
            }
        }
    }
}
