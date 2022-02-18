using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.Drawing;
using System.IO;

namespace Format_and_rotate_image
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
                IWParagraph paragraph = section.AddParagraph();
                paragraph.AppendText("This paragraph has picture. ");
                FileStream imageStream = new FileStream(Path.GetFullPath(@"../../../Image.png"), FileMode.Open, FileAccess.ReadWrite);
                //Appends new picture to the paragraph.
                WPicture picture = paragraph.AppendPicture(imageStream) as WPicture;
                //Sets text wrapping style – When the wrapping style is inline, the images are not absolutely positioned. It is added next to the text range.
                picture.TextWrappingStyle = TextWrappingStyle.Square;
                //Sets horizontal and vertical origin.
                picture.HorizontalOrigin = HorizontalOrigin.Page;
                picture.VerticalOrigin = VerticalOrigin.Paragraph;
                //Sets width and height for the paragraph.
                picture.Width = 150;
                picture.Height = 100;
                //Sets horizontal and vertical position for the picture.
                picture.HorizontalPosition = 200;
                picture.VerticalPosition = 150;
                //Sets lock aspect ratio for the picture.
                picture.LockAspectRatio = true;
                picture.Name = "PictureName";
                //Sets horizontal and vertical alignments.
                picture.HorizontalAlignment = ShapeHorizontalAlignment.Center;
                picture.VerticalAlignment = ShapeVerticalAlignment.Bottom;
                //Sets 90 degree rotation.
                picture.Rotation = 90;
                //Sets horizontal flip.
                picture.FlipHorizontal = true;
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
