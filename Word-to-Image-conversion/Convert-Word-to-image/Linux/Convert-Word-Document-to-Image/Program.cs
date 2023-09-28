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
            using (FileStream docStream = new FileStream("Data/Input.docx", FileMode.Open, FileAccess.Read))
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
                        //Create FileStream to save the image file.
                        using (FileStream fileStreamOutput = new FileStream("WordToImage.Jpeg", FileMode.Create, FileAccess.ReadWrite, FileShare.ReadWrite))
                        {
                            imageStream.CopyTo(fileStreamOutput);
                        }
                    }
                }
            }
        }
    }
}
