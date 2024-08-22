using System.IO;
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.DocIORenderer;

namespace First_page_of_Word_to_image
{
    class Program
    {
        static void Main(string[] args)
        {
            using (FileStream fileStream = new FileStream(Path.GetFullPath(@"Data/Template.docx"), FileMode.Open))
            {
                //Loads an existing Word document.
                using (WordDocument wordDocument = new WordDocument(fileStream, FormatType.Automatic))
                {
                    //Creates an instance of DocIORenderer.
                    using (DocIORenderer renderer = new DocIORenderer())
                    {
                        //Convert the first page of the Word document into an image.
                        Stream imageStream = wordDocument.RenderAsImages(0, ExportImageFormat.Jpeg);
                        //Resets the stream position.
                        imageStream.Position = 0;
                        //Creates the output image file stream.
                        using (FileStream fileStreamOutput = File.Create(Path.GetFullPath(@"Output/Output.jpeg")))
                        {
                            //Copies the converted image stream into created output stream.
                            imageStream.CopyTo(fileStreamOutput);
                        }
                    }
                }
            }
        }
    }
}
