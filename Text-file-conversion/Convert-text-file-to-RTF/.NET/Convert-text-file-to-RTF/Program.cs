using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;

namespace Convert_text_file_to_RTF
{
    class Program
    {
        static void Main(string[] args)
        {
            //Loads an existing Word document into DocIO instance.
            using (FileStream fileStreamPath = new FileStream(Path.GetFullPath(@"Data/Template.txt"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                using (WordDocument document = new WordDocument(fileStreamPath, FormatType.Txt))
                {
                    //Creates file stream.
                    using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"Output/Result.rtf"), FileMode.Create, FileAccess.ReadWrite))
                    {
                        //Saves the Word document to file stream.
                        document.Save(outputFileStream, FormatType.Rtf);
                    }
                }
            }
        }
    }
}
