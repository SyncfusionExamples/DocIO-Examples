using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;

namespace Remove_section_from_document
{
    class Program
    {
        static void Main(string[] args)
        {
            using (FileStream fileStreamPath = new FileStream(Path.GetFullPath(@"Data/Template.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                //Opens an input Word template.
                using (WordDocument document = new WordDocument(fileStreamPath, FormatType.Automatic))
                {
                    //Removes the second section from the collection.
                    document.Sections.RemoveAt(1);
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
