using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.Drawing;
using System.IO;

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
                document.Open("Input.md");

                //Create a file stream.
                using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"../../../Sample.docx"), FileMode.Create, FileAccess.ReadWrite))
                {
                    //Save a Markdown file to the file stream.
                    document.Save(outputFileStream, FormatType.Docx);
                }
            }
        }
        private static void MdImportSettings_ImageNodeVisited(object sender, Syncfusion.Office.Markdown.MdImageNodeVisitedEventArgs args)
        {
            //Set image stream based on image name from input Markdown.
            if (args.Uri == "Image_1.png")
                args.ImageStream = new FileStream("Image_1.png", FileMode.Open);
            else if (args.Uri == "Image_2.png")
                args.ImageStream = new FileStream("Image_2.png", FileMode.Open);
            //Retrive image from website and use
            else if (args.Uri.StartsWith("https://"))
            {
                WebClient client = new WebClient();
                //Download the image as stream.
                byte[] image = client.DownloadData(args.Uri);
                Stream stream = new MemoryStream(image);
                //Set the retrieved image from the input Markdown.
                args.ImageStream = stream;
            }
        }
    }
}