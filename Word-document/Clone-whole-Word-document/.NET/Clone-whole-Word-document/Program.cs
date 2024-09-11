using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System;
using System.IO;

namespace Clone_whole_Word_document
{
    class Program
    {
        static void Main(string[] args)
        {
            using (FileStream fileStreamPath = new FileStream(Path.GetFullPath(@"Data/Template.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                //Opens an existing document from file system through constructor of WordDocument class.
                using (WordDocument document = new WordDocument(fileStreamPath, FormatType.Automatic))
                {
                    //Creates a clone of Input Template.
                    using (WordDocument clonedDocument = document.Clone())
                    {
                        //Creates file stream.
                        using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"Output/Result.docx"), FileMode.Create, FileAccess.ReadWrite))
                        {
                            //Saves the cloned document instance.
                            clonedDocument.Save(outputFileStream, FormatType.Docx);
                        }
                    }
                }
            }
        }
    }
}
