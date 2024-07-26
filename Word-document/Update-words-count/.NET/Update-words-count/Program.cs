using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.DocIORenderer;
using System.IO;

namespace Update_words_count
{
    class Program
    {
        static void Main(string[] args)
        {
            using (FileStream fileStreamPath = new FileStream(Path.GetFullPath(@"../../../Template.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                //Opens an document from file system through constructor of WordDocument class.
                using (WordDocument document = new WordDocument(fileStreamPath, FormatType.Automatic))
                {
                    //Update the word count in the document.
                    document.UpdateWordCount(false);
                    //Get the word count in the document.
                    int wordCount = document.BuiltinDocumentProperties.WordCount;
                    //Get the character count in the document.
                    int charCount = document.BuiltinDocumentProperties.CharCount;
                    //Get the paragraph count in the document.
                    int paragraphCount = document.BuiltinDocumentProperties.ParagraphCount;
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
