using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;

namespace Add_image
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
                //Adds image to the paragraph.
                FileStream imageStream = new FileStream(Path.GetFullPath(@"../../../Image.png"), FileMode.Open, FileAccess.ReadWrite);
                IWPicture picture = firstParagraph.AppendPicture(imageStream);
                //Sets height and width for the image.
                picture.Height = 100;
                picture.Width = 200;
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
