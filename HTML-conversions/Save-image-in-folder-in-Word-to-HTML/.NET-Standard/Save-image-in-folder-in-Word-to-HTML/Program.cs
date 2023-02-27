using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;

namespace Save_image_in_folder_in_Word_to_HTML
{
    class Program
    {
        static void Main(string[] args)
        {
            //Load an existing Word document. 
            using (FileStream fileStreamPath = new FileStream(Path.GetFullPath(@"../../../Input.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                //Load the file stream into a Word document.
                using (WordDocument document = new WordDocument(fileStreamPath, FormatType.Docx))
                {
                    //Create a file stream.
                    using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"../../../WordToHtml.html"), FileMode.Create, FileAccess.ReadWrite))
                    {
                        //Hook the event to customize the image. 
                        document.SaveOptions.ImageNodeVisited += SaveImage;
                        //Save the html file to the file stream.
                        document.Save(outputFileStream, FormatType.Html);
                       
                    }                                      
                }
            }
        }
        static int imageCount = 0;
        static void SaveImage(object sender, ImageNodeVisitedEventArgs args)
        {
            //Customize the image path and save the image in an external folder.
            string imagepath = @"D:\Temp\Image_" + imageCount + ".png";
            //Save the image stream as a file.
            using (FileStream fileStreamOutput = File.Create(imagepath))
                args.ImageStream.CopyTo(fileStreamOutput);
            //Set the image URI to be used in the output HTML.
            args.Uri = imagepath;
            imageCount++;
        }
    }
}
