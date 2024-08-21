using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;

namespace Plain_text_content_control
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
                //Appends plain text content control to the paragraph.
                InlineContentControl plainTextControl = paragraph.AppendInlineContentControl(ContentControlType.Text) as InlineContentControl;
                WTextRange textRange = new WTextRange(document);
                textRange.Text = "Plain text content control.";
                //Adds new text to the plain text content control.
                plainTextControl.ParagraphItems.Add(textRange);
                //Creates file stream.
                using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"Output/Output.docx"), FileMode.Create, FileAccess.ReadWrite))
                {
                    //Saves the Word document to file stream.
                    document.Save(outputFileStream, FormatType.Docx);
                }
            }
        }
    }
}
