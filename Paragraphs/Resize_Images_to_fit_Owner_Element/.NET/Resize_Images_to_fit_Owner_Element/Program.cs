using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;

namespace Resize_Images_to_fit_Owner_Element
{
    internal class Program
    {
        static void Main(string[] args)
        {
            // Create a new Word document
            using (WordDocument document = new WordDocument())
            {
                // Enable automatic image resizing
                document.Settings.ResizeImageToFitInContainer = true;
                // Open the template document
                using (FileStream fileStreamPath = new FileStream(Path.GetFullPath(@"Data/InputDocument.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    document.Open(fileStreamPath, FormatType.Automatic);
                }
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