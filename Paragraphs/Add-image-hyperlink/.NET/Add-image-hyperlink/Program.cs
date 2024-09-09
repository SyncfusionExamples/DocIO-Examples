using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;

namespace Add_image_hyperlink
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
                paragraph.AppendText("Image Hyperlink");
                paragraph = section.AddParagraph();
                //Creates a new image instance and load image.
                WPicture picture = new WPicture(document);
                FileStream imageStream = new FileStream(Path.GetFullPath(@"Data/Mountain-200.jpg"), FileMode.Open, FileAccess.ReadWrite);
                picture.LoadImage(imageStream);
                //Appends new image hyperlink to the paragraph.
                paragraph.AppendHyperlink("http://www.syncfusion.com", picture, HyperlinkType.WebLink);
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
