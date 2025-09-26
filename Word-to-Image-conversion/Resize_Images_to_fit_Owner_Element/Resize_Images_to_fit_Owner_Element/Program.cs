using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;

namespace Resize_Images_to_fit_Owner_Element
{
    internal class Program
    {
        static void Main(string[] args)
        {
            // Create a new Word document instance
            WordDocument document = new WordDocument();
            // Enable automatic image resizing
            document.Settings.ResizeImageToFitInContainer = true;
            // Open the template document
            document.Open(Path.GetFullPath(@"Data/InputDocument.docx"), FormatType.Docx);
            // Save the modified document
            document.Save(Path.GetFullPath(@"Output/Result.docx"), FormatType.Docx);
            // Release resources
            document.Dispose();
        }
    }
}