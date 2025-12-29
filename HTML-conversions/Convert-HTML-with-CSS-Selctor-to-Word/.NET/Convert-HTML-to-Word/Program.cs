using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;

namespace Convert_HTML_to_Word
{
    class Program
    {
        static void Main(string[] args)
        {
            //Loads an existing Word document into DocIO instance.
            using (FileStream fileStreamPath = new FileStream(Path.GetFullPath(@"Data/Input.html"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                using (WordDocument document = new WordDocument(fileStreamPath, FormatType.Html))
                {
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
}
