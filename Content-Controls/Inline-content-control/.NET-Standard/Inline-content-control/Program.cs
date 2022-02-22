using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;

namespace Inline_content_control
{
    class Program
    {
        static void Main(string[] args)
        {
            //Creates a new Word document.
            using (WordDocument document = new WordDocument())
            {
                //Adds one section and one paragraph to the document.
                document.EnsureMinimal();
                //Gets the last paragraph.
                WParagraph paragraph = document.LastParagraph;
                //Adds text to the paragraph.
                paragraph.AppendText("A new text is added to the paragraph. ");
                //Appends inline content control to the paragraph.
                InlineContentControl inlineContentControl = paragraph.AppendInlineContentControl(ContentControlType.RichText) as InlineContentControl;
                WTextRange textRange = new WTextRange(document);
                textRange.Text = "Inline content control ";
                //Adds new text to the inline content control.
                inlineContentControl.ParagraphItems.Add(textRange);
                //Creates file stream.
                using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"../../../Result.docx"), FileMode.Create, FileAccess.ReadWrite))
                {
                    //Saves the Word document to file stream.
                    document.Save(outputFileStream, FormatType.Docx);
                }
            }
        }
    }
}
