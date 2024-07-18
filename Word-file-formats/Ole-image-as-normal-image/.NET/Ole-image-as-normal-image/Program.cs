using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;

namespace Ole_image_as_normal_image
{
    class Program
    {
        static void Main(string[] args)
        {
            //Creates a new instance of WordDocument (Empty Word Document).
            using (WordDocument document = new WordDocument())
            {
                //Loads or opens an existing Word document from stream.
                using (FileStream fileStreamPath = new FileStream(Path.GetFullPath(@"../../../Template.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    //Sets flag to preserve embedded Ole image as normal image while opening document.
                    document.Settings.PreserveOleImageAsImage = true;
                    //Loads or opens an existing Word document through Open method of WordDocument class. 
                    document.Open(fileStreamPath, FormatType.Automatic);
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
