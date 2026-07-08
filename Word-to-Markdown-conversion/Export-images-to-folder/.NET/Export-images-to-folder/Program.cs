using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.Office;
using System;
using System.IO;

namespace Export_images_to_folder
{
    class Program
    {
        static void Main(string[] args)
        {
            //Open an existing Word document.
            using (WordDocument document = new WordDocument(Path.GetFullPath(@"Data/Input.docx")))
            {
                //Set images folder to export images. 
                document.SaveOptions.MarkdownExportImagesFolder = Path.GetFullPath(@"Output/");
                //Save the document as a Markdown file.
                document.Save(Path.GetFullPath(@"Output/Output.md"));  
            }
        }
    }
}
