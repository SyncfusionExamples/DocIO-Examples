using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;

namespace Apply_picture_background_to_document
{
    class Program
    {
        static void Main(string[] args)
        {
            using (FileStream fileStreamPath = new FileStream(Path.GetFullPath(@"../../../Data/Template.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                //Opens an document from file system through constructor of WordDocument class.
                using (WordDocument document = new WordDocument(fileStreamPath, FormatType.Automatic))
                {
                    //Sets the background type as picture.
                    document.Background.Type = BackgroundType.Picture;
                    //Opens the existing image. 
                    using (FileStream imageStream = new FileStream(@"../../../Data/Picture.png", FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                    {
                        using (MemoryStream memoryStream = new MemoryStream())
                        {
                            imageStream.CopyTo(memoryStream);
                            document.Background.Picture = memoryStream.ToArray();
                        }
                    }
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
