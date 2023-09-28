using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.DocIORenderer;
using System.IO;

namespace Convert_Word_Document_to_Image.Data
{
    public class WordService
    {
        public  MemoryStream ConvertWordtoImage()
        {
            //Open the file as Stream
            using (FileStream sourceStreamPath = new FileStream(@"wwwroot/Input.docx", FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                //Open an existing Word document.
                using (WordDocument document = new WordDocument(sourceStreamPath, FormatType.Docx))
                {
                    //Instantiation of DocIORenderer for Word to Image conversion
                    using (DocIORenderer render = new DocIORenderer())
                    {
                        Stream imageStream = document.RenderAsImages(0, ExportImageFormat.Jpeg);
                        //Reset the stream position.
                        imageStream.Position = 0;
                        return (MemoryStream)imageStream;
                    }
                }
            }
        }
    }
}
