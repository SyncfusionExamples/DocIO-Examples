using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;

namespace Convert_WordML_to_Word
{
    class Program
    {
        static void Main(string[] args)
        {
            //Loads an existing Word document into DocIO instance.
            using (FileStream fileStreamPath = new FileStream(Path.GetFullPath(@"../../../Template.xml"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                using (WordDocument document = new WordDocument(fileStreamPath, FormatType.WordML))
                {
                    //Creates file stream.
                    using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"../../../WordMLToWord.docx"), FileMode.Create, FileAccess.ReadWrite))
                    {
                        //Saves the Word document to file stream.
                        document.Save(outputFileStream, FormatType.Docx);
                    }
                }
            }
        }
    }
}
