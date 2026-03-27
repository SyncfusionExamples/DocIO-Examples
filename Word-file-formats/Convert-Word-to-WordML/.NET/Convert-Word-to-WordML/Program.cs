using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;

namespace Convert_Word_to_WordML
{
    class Program
    {
        static void Main(string[] args)
        {
            //Loads an existing Word document into DocIO instance
            using (FileStream fileStreamPath = new FileStream(Path.GetFullPath(@"Data/Template.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                using (WordDocument document = new WordDocument(fileStreamPath, FormatType.Docx))
                {
                    //Creates file stream
                    using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"Output/Result.xml"), FileMode.Create, FileAccess.ReadWrite))
                    {
                        //Save the loaded document in WordML format to the output stream
                        document.Save(outputFileStream, FormatType.WordML);
                        //Closes the Word document
                        document.Close();
                    }
                }
            }
        }      
    }
}
