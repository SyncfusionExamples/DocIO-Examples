using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;

namespace Remove_Header_Footer_In_First_Page
{
    class Program
    {
        static void Main(string[] args)
        {
            //Creates a new Word document.
            using (FileStream fileStreamPath = new FileStream(@"../../../Data/Template.docx", FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                using (WordDocument document = new WordDocument(fileStreamPath, Syncfusion.DocIO.FormatType.Automatic))
                {
                    //Get the first section from the Word document
                    WSection section = document.Sections[0];
                    //Set DifferentFirstPage as true for indicating different header and footer used on first page.
                    section.PageSetup.DifferentFirstPage = true;
                    using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"../../../Output/Result.docx"), FileMode.Create, FileAccess.ReadWrite))
                    {
                        //Save a Word document to the file stream.
                        document.Save(outputFileStream, Syncfusion.DocIO.FormatType.Docx);
                    }
                }
            }
        }
    }
}