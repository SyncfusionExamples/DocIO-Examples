using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;

namespace Convert_Word_to_Markdown.Data
{
    public class WordService
    {
        public MemoryStream ConvertWordToMD()
        {
            using (FileStream inputStream = new FileStream(@"wwwroot/sample-data/Input.docx", FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                //Open an existing Word document.
                using (WordDocument document = new WordDocument(inputStream, FormatType.Automatic))
                {
                    //Save as a Markdown file into the MemoryStream.
                    using (MemoryStream stream = new MemoryStream())
                    {
                        document.Save(stream, FormatType.Markdown);
                        stream.Position = 0;
                        return stream;
                    }
                }
            }
        }
    }
}
