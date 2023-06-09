using System;
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;
using static System.Collections.Specialized.BitVector32;
using Syncfusion.DocIORenderer;


namespace Convert_Word_Document_to_Image
{
    class Program
    {
        static void Main(string[] args)
        {
            //Open the file as Stream
            using (FileStream docStream = new FileStream(Path.GetFullPath(@"../../../Data/Input.docx"), FileMode.Open, FileAccess.Read))
            {
                //Loads an existing  Word document
                using (WordDocument wordDocument = new WordDocument(docStream, FormatType.Docx))
                {
                    //Instantiation of DocIORenderer for Word to image conversion
                    using (DocIORenderer render = new DocIORenderer())
                    {
                        //Convert the first page of the Word document into an image.
                        Stream imageStream = wordDocument.RenderAsImages(0, ExportImageFormat.Jpeg);
                        //Reset the stream position.
                        imageStream.Position = 0;
                        //Save the stream as file.
                        using (FileStream fileStreamOutput = File.Create("wordtoimage.jpeg"))
                        {
                            imageStream.CopyTo(fileStreamOutput);
                        }
                    }
                }
            }
        }
    }
}
