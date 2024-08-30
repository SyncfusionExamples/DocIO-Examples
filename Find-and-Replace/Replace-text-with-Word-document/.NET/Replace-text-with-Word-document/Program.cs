using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.Drawing;
using System.IO;

namespace Replace_text_with_Word_document
{
    class Program
    {
        static void Main(string[] args)
        {
            //Loads a template document.
            using (FileStream fileStreamPath = new FileStream(Path.GetFullPath(@"Data/SourceTemplate.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                using (WordDocument document = new WordDocument(fileStreamPath, FormatType.Docx))
                {
                    //Gets the document to replace the text.
                    using (FileStream replaceFileStreamPath = new FileStream(Path.GetFullPath(@"Data/ReplacementDoc.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                    {
                        //Opens an existing Word document.
                        using (IWordDocument replaceDocument = new WordDocument(replaceFileStreamPath, FormatType.Docx))
                        {
                            //Replaces a particular text with another document.
                            document.Replace("Information", replaceDocument, false, true, true);
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
    }
}
