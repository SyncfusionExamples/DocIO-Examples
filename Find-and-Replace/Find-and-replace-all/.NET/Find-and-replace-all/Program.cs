using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System;
using System.IO;

namespace Find_and_replace_all
{
    class Program
    {
        static void Main(string[] args)
        {
            using (FileStream fileStream = new FileStream(Path.GetFullPath(@"../../../Giant Panda.docx"), FileMode.Open, FileAccess.ReadWrite))
            {
                //Loads an existing Word document into DocIO instance.
                using (WordDocument document = new WordDocument(fileStream, FormatType.Automatic))
                {
                    //Replaces the word "bear" as "panda".
                    document.Replace("bear", "panda", false, true);
                    //Creates file stream.
                    using (FileStream outputStream = new FileStream(Path.GetFullPath(@"../../../Result.docx"), FileMode.Create, FileAccess.ReadWrite))
                    {
                        //Saves the Word document to file stream.
                        document.Save(outputStream, FormatType.Docx);
                    }
                }
            }
        }
    }
}
