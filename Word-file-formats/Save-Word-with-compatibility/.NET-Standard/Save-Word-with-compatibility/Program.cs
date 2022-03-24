using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;

namespace Save_Word_with_compatibility
{
    class Program
    {
        static void Main(string[] args)
        {
            //Creates an empty WordDocument instance.
            using (WordDocument document = new WordDocument())
            {
                //Loads or opens an existing Word document from stream.
                using (FileStream fileStreamPath = new FileStream(Path.GetFullPath(@"../../../Template.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    //Loads or opens an existing Word document through Open method of WordDocument class.
                    document.Open(fileStreamPath, FormatType.Automatic);
                    //Enables flag to maintain compatibility with same Word version.
                    document.SaveOptions.MaintainCompatibilityMode = true;
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
