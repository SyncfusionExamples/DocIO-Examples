using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.Drawing;
using System.IO;

namespace Customize_image_path
{
    class Program
    {
        static void Main(string[] args)
        {
            //Load the file stream into a Word document.
            using (WordDocument document = new WordDocument(Path.GetFullPath(@"Data/Input.docx")))
            {
                //Hook the event to customize the image. 
                document.SaveOptions.MarkdownSaveOptions.ImageNodeVisited += SaveImage;
                //Save a Markdown file to the file stream.
                document.Save(Path.GetFullPath(@"Output/Output.md"));
            }
        }
        //The following code examples show the event handler to customize the image path and save the image in an external folder.
        static void SaveImage(object sender, ImageNodeVisitedEventArgs args)
        {
            string imagepath = Path.GetFullPath(@"Output/Output.png");
            //Save the image stream as a file. 
            using (FileStream fileStreamOutput = File.Create(imagepath))
                args.ImageStream.CopyTo(fileStreamOutput);
            //Set the image URI to be used in the output markdown.
            args.Uri = imagepath;
        }
    }
}
