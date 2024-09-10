using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;

namespace Add_text_box
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
                IWTextBox textbox = paragraph.AppendTextBox(150, 115);
                //Adds new text to the textbox body.
                IWParagraph textboxParagraph = textbox.TextBoxBody.AddParagraph();
                textboxParagraph.AppendText("Text inside text box");
                textboxParagraph = textbox.TextBoxBody.AddParagraph();
                //Adds new picture to textbox body.
                FileStream imagestream = new FileStream(Path.GetFullPath(@"Data/Mountain-200.jpg"), FileMode.Open, FileAccess.ReadWrite);
                IWPicture picture = textboxParagraph.AppendPicture(imagestream);
                picture.Height = 90;
                picture.Width = 110;
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
