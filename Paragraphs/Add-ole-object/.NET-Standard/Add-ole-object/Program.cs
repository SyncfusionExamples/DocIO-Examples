using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;

namespace Add_ole_object
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
                //Opens the file to be embedded.
                using (FileStream fileStream = new FileStream(Path.GetFullPath(@"../../../Data/Book.xlsx"), FileMode.Open, FileAccess.ReadWrite))
                {
                    //Loads the picture instance with the image need to be displayed.
                    WPicture picture = new WPicture(document);
                    FileStream imageStream = new FileStream(Path.GetFullPath(@"../../../Data/Image.png"), FileMode.Open, FileAccess.ReadWrite);
                    picture.LoadImage(imageStream);
                    //Appends the OLE object to the paragraph.
                    WOleObject oleObject = paragraph.AppendOleObject(fileStream, picture, OleObjectType.ExcelWorksheet);
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
}
