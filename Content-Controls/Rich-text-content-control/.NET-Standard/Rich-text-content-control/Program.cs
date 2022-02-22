using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;

namespace Rich_text_content_control
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
                //Appends rich text content control to the paragraph.
                InlineContentControl richTextControl = paragraph.AppendInlineContentControl(ContentControlType.RichText) as InlineContentControl;
                WTextRange textRange = new WTextRange(document);
                textRange.Text = "Rich text content control.";
                //Adds new text to the rich text content control.
                richTextControl.ParagraphItems.Add(textRange);
                WPicture picture = new WPicture(document);
                Stream imageStream = new FileStream(Path.GetFullPath(@"../../../Image.png"), FileMode.Open, FileAccess.Read);
                //Adds image from stream.
                picture.LoadImage(imageStream);
                picture.Height = 100;
                picture.Width = 100;
                //Adds new picture to the rich text content control.
                richTextControl.ParagraphItems.Add(picture);
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
