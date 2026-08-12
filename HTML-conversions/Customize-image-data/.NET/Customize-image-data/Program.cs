using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;

namespace Customize_image_data
{
    class Program
    {
        static void Main(string[] args)
        { 
            //Creates a new instance of WordDocument
            using (WordDocument document = new WordDocument())
            {
                //Hooks the ImageNodeVisited event to open the image from a specific location
                document.HTMLImportSettings.ImageNodeVisited += OpenImage;
                //Opens the input HTML document
                document.Open(@"Data\Input.html", FormatType.Html);
                //Unhooks the ImageNodeVisited event after loading HTML
                document.HTMLImportSettings.ImageNodeVisited -= OpenImage;
                //Saves the Word document
                document.Save(@"Output\HtmlToWord.docx", FormatType.Docx);
                //Closes the WordDocument instance
                document.Close();
            }
        }
        private static void OpenImage(object sender, ImageNodeVisitedEventArgs args)
        {
            //Read the image from the specified (args.Uri) path.
            args.ImageStream = System.IO.File.OpenRead(args.Uri);
        }
    }

}
