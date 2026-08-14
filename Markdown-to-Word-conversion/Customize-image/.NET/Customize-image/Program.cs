using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.Drawing;
using System.IO;
using System.Net;
using System.Net.Http;

namespace Customize_image
{
    class Program
    {
        static void Main(string[] args)
        {
            //Create a Word document instance.
            using (WordDocument document = new WordDocument())
            {
                //Hook the event to customize the image while importing Markdown.
                document.MdImportSettings.ImageNodeVisited += MdImportSettings_ImageNodeVisited;
                //Open the Markdown file.
                document.Open(Path.GetFullPath("Data/Input.md"));
                //Save as a Word document.
                document.Save(Path.GetFullPath(@"../../../Output/Output.docx"));
            }
        }
        private static void MdImportSettings_ImageNodeVisited(object sender, Syncfusion.Office.Markdown.MdImageNodeVisitedEventArgs args)
        {
            //Set the image stream based on the image name from the input Markdown.
            if (args.Uri == "Image_1.png")
                args.ImageStream = new FileStream(Path.GetFullPath("Data/Image_1.png"), FileMode.Open);
            else if (args.Uri == "Image_2.png")
                args.ImageStream = new FileStream(Path.GetFullPath("Data/Image_2.png"), FileMode.Open);
            //Retrive the image from the website and use it.
            else if (args.Uri.StartsWith("https://"))
            {
                //Download the image as a stream.
                using (HttpClient client = new HttpClient())
                {
                    byte[] image = client.GetByteArrayAsync(args.Uri).GetAwaiter().GetResult();
                    args.ImageStream = new MemoryStream(image);
                }
            }
        }
    }
}