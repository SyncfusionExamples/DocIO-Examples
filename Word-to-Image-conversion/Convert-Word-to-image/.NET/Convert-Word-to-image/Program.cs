using System.IO;
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.DocIORenderer;

namespace Convert_Word_to_image
{
    class Program
    {
        static void Main(string[] args)
        {
            // Open the Word document file stream
            using (FileStream fileStream = new FileStream(Path.GetFullPath(@"Data/Template.docx"), FileMode.Open))
            {
                //Loads an existing Word document.
                using (WordDocument wordDocument = new WordDocument(fileStream, FormatType.Automatic))
                {
                    //Creates an instance of DocIORenderer.
                    using (DocIORenderer renderer = new DocIORenderer())
                    {
                        //Convert the entire Word document to images.
                        Stream[] imageStreams = wordDocument.RenderAsImages();
                        for (int i = 0; i < imageStreams.Length; i++)
                        {
                            //Save the image stream as file.
                            using (FileStream fileStreamOutput = File.Create(Path.GetFullPath(@"Output/Output_" + i + ".jpeg")))
                            {
                                imageStreams[i].CopyTo(fileStreamOutput);
                            }
                        }
                    }
                }
            }
        }
    }
}
