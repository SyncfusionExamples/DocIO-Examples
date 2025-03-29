using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.Drawing;
using System.IO;

namespace Apply_image_border
{
    class Program
    {
        static void Main(string[] args)
        {
            //Create an word document.
            using (WordDocument document = new WordDocument())
            {
                document.EnsureMinimal();
                //Appends new textbox to the paragraph
                IWTextBox textbox = document.LastParagraph.AppendTextBox(150, 75);
                //Set color to the text box's line.
                textbox.TextBoxFormat.LineColor = Color.Purple;
                //Set size of the text box's border.
                textbox.TextBoxFormat.LineWidth = 2;
                //Sets text box's margin values as Zero.
                textbox.TextBoxFormat.InternalMargin.Top = 0f;
                textbox.TextBoxFormat.InternalMargin.Bottom = 0f;
                textbox.TextBoxFormat.InternalMargin.Left = 0f;
                textbox.TextBoxFormat.InternalMargin.Right = 0f;
                //Set text trapping style to the text box.
                textbox.TextBoxFormat.TextWrappingStyle = TextWrappingStyle.Inline;

                //Adds new text to the textbox body
                IWParagraph textboxParagraph = textbox.TextBoxBody.AddParagraph();
                FileStream imageStream = new FileStream(Path.GetFullPath(@"Data/Picture.png"), FileMode.Open, FileAccess.ReadWrite);
                WPicture picture = textboxParagraph.AppendPicture(imageStream) as WPicture;
                //sets the picture width scale factor in percent.
                picture.WidthScale = 80;
                //sets the picture height scale factor in percent.
                picture.HeightScale = 80;

                //Set picture size as text box size.
                textbox.TextBoxFormat.Width = picture.Width;
                textbox.TextBoxFormat.Height = picture.Height;

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

