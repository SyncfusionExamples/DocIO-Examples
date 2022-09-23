using System.IO;
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.DocIORenderer;

namespace Specific_range_of_pages_Word_to_image
{
    class Program
    {
        static void Main(string[] args)
        {
            using (FileStream fileStream = new FileStream(Path.GetFullPath(@"../../../Template.docx"), FileMode.Open))
            {
                //Loads an existing Word document.
                using (WordDocument wordDocument = new WordDocument(fileStream, FormatType.Automatic))
                {
                    //Creates an instance of DocIORenderer.
                    using (DocIORenderer renderer = new DocIORenderer())
                    {
                        //Convert a specific range of pages in Word document to images.
                        Stream[] imageStreams = wordDocument.RenderAsImages(1, 2);
                        int i = 0;
                        foreach (Stream stream in imageStreams)
                        {
                            //Resets the stream position.
                            stream.Position = 0;
                            //Creates the output image file stream.
                            using (FileStream fileStreamOutput = File.Create(Path.GetFullPath(@"../../../WordToImage_" + i + ".jpeg")))
                            {
                                //Copies the converted image stream into created output stream.
                                stream.CopyTo(fileStreamOutput);
                            }
                            i++;
                        }
                    }
                }
            }
        }
    }
}
