using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.Drawing;
using System.IO;

namespace Format_and_rotate_text_box
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
                //Appends new textbox to the paragraph.
                IWTextBox textbox = paragraph.AppendTextBox(150, 75);
                //Adds new text to the textbox body.
                IWParagraph textboxParagraph = textbox.TextBoxBody.AddParagraph();
                textboxParagraph.AppendText("Text inside text box");
                //Sets fill color and line width for textbox.
                textbox.TextBoxFormat.FillColor = Color.LightGreen;
                textbox.TextBoxFormat.LineWidth = 2;
                //Applies textbox text direction.
                textbox.TextBoxFormat.TextDirection = Syncfusion.DocIO.DLS.TextDirection.VerticalTopToBottom;
                //Sets text wrapping style.
                textbox.TextBoxFormat.TextWrappingStyle = TextWrappingStyle.InFrontOfText;
                //Sets horizontal and vertical position.
                textbox.TextBoxFormat.HorizontalPosition = 200;
                textbox.TextBoxFormat.VerticalPosition = 200;
                //Sets horizontal and vertical origin.
                textbox.TextBoxFormat.VerticalOrigin = VerticalOrigin.Margin;
                textbox.TextBoxFormat.HorizontalOrigin = HorizontalOrigin.Page;
                //Sets top and bottom margin values.
                textbox.TextBoxFormat.InternalMargin.Bottom = 5f;
                textbox.TextBoxFormat.InternalMargin.Top = 5f;
                //Sets 90 degree rotation.
                textbox.TextBoxFormat.Rotation = 90;
                //Sets horizontal flip.
                textbox.TextBoxFormat.FlipHorizontal = true;
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
