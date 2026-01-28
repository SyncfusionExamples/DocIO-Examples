using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.DocIORenderer;
using System.Drawing;

namespace Convert_Word_document_to_Thumbnail_Image
{
    class Program
    {
        static void Main(string[] args)
        {
            //Open the Word document file stream. 
            using (FileStream fileStream = new FileStream(Path.GetFullPath(@"Data/Input.docx"), FileMode.Open))
            {
                //Loads an existing Word document.
                using (WordDocument wordDocument = new WordDocument(fileStream, FormatType.Automatic))
                {
                    //Creates an instance of DocIORenderer.
                    using (DocIORenderer renderer = new DocIORenderer())
                    {
                        //Convert the first page of the Word document into an image.
                        Stream imageStream = wordDocument.RenderAsImages(0, ExportImageFormat.Png);
                        //Reset the stream position.
                        imageStream.Position = 0;

                        //Resize image to thumbnail size.
                        Image image = Image.FromStream(imageStream);
                        Image thumbnail = image.GetThumbnailImage(600, 700, () => false, IntPtr.Zero);

                        //Save the image.
                        thumbnail.Save(Path.GetFullPath(@"Output/Image1.png"), System.Drawing.Imaging.ImageFormat.Png);
                        thumbnail.Dispose();
                        imageStream.Dispose();
                        image.Dispose();
                    }
                }
            }
        }
    }
}