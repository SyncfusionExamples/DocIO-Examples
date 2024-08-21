using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;

namespace Picture_content_control
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
                //Appends picture content control to the paragraph.
                InlineContentControl pictureContentControl = paragraph.AppendInlineContentControl(ContentControlType.Picture) as InlineContentControl;
                //Creates a new image instance and load image.
                WPicture picture = new WPicture(document);
                Stream imageStream = new FileStream(Path.GetFullPath(@"Data/Image.png"), FileMode.Open, FileAccess.Read);
                //Adds image from stream.
                picture.LoadImage(imageStream);
                //Adds picture to the picture content control.
                pictureContentControl.ParagraphItems.Add(picture);
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
