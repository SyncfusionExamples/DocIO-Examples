
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;


namespace Customize_image_path_in_Word_to_html
{
    class Program
    {
        static void Main(string[] args)
        {
            //Open the file as a Stream.
            using (FileStream docStream = new FileStream(Path.GetFullPath(@"Data/Input.docx"), FileMode.Open, FileAccess.Read))
            {
                //Load the file stream into a Word document.
                using (WordDocument document = new WordDocument(docStream, FormatType.Docx))
                {
                    //Hook the event to customize the image. 
                    document.SaveOptions.ImageNodeVisited += SaveImage;
                    using (FileStream outputStream = new FileStream(Path.GetFullPath(@"Output/WordToHTML.html"), FileMode.Create, FileAccess.ReadWrite, FileShare.ReadWrite))
                    {
                        //Save the HTML file.
                        document.Save(outputStream, FormatType.Html);
                    }
                }
            }
            static void SaveImage(object sender, ImageNodeVisitedEventArgs args)
            {
                // Specify the folder path.
                string folderPath = Path.GetFullPath(@"Output");
                // Specify the image filename.
                string imageFilename = "Image.png"; 
                // Check if the folder exists, and create it if it doesn't.
                if (!Directory.Exists(folderPath))
                {
                    Directory.CreateDirectory(folderPath);
                }
                string imageFullPath = Path.Combine(folderPath, imageFilename);
                // Save the image stream as a file.
                using (FileStream fileStreamOutput = File.Create(imageFullPath))
                {
                    args.ImageStream.CopyTo(fileStreamOutput);
                }
                // Set the URI to be used for the image in the output HTML.
                args.Uri = imageFullPath;
            }
        }
    }
}
