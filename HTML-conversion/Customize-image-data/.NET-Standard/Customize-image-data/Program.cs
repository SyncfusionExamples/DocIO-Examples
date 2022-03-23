using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;

namespace Customize_image_data
{
    class Program
    {
        static void Main(string[] args)
        {
            //Loads an existing Word document into DocIO instance. 
            using (FileStream fileStreamPath = new FileStream(Path.GetFullPath(@"../../../Data/Input.html"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                using (WordDocument document = new WordDocument())
                {
                    //Hooks the ImageNodeVisited event to open the image from a specific location.
                    document.HTMLImportSettings.ImageNodeVisited += OpenImage;
                    //Opens the input HTML document.
                    document.Open(fileStreamPath, FormatType.Html);
                    //Unhooks the ImageNodeVisited event after loading HTML.
                    document.HTMLImportSettings.ImageNodeVisited -= OpenImage;
                    //Creates file stream.
                    using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"../../../HtmlToWord.docx"), FileMode.Create, FileAccess.ReadWrite))
                    {
                        //Saves the Word document to file stream.
                        document.Save(outputFileStream, FormatType.Docx);
                    }
                }
            }
        }
        private static void OpenImage(object sender, ImageNodeVisitedEventArgs args)
        {
            //Read the image from the specified (args.Uri) path.
            args.ImageStream = System.IO.File.OpenRead(args.Uri);
        }
    }
}
