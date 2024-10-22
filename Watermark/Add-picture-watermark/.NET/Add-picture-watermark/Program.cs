using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;

namespace Add_picture_watermark
{
    class Program
    {
        static void Main(string[] args)
        {
            //Creates a new Word document.
            using (WordDocument document = new WordDocument())
            {
                //Adds a section and a paragraph in the document.
                document.EnsureMinimal();
                IWParagraph paragraph = document.LastParagraph;
                paragraph.AppendText("The Northwind sample database (Northwind.mdb) is included with all versions of Access. It provides data you can experiment with and database objects that demonstrate features you might want to implement in your own databases.");
                //Creates a new picture watermark.
                PictureWatermark picWatermark = new PictureWatermark();
                //Sets the scaling to picture.
                picWatermark.Scaling = 120f;
                picWatermark.Washout = true;
                //Sets the picture watermark to document.
                document.Watermark = picWatermark;
                FileStream imageStream = new FileStream(Path.GetFullPath(@"Data/Northwind-logo.png"), FileMode.Open, FileAccess.Read);
                BinaryReader br = new BinaryReader(imageStream);
                byte[] image = br.ReadBytes((int)imageStream.Length);
                //Sets the image to the picture watermark.
                picWatermark.LoadPicture(image);
                //Creates file stream.
                using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"Output/Result.docx"), FileMode.Create, FileAccess.ReadWrite))
                {
                    //Saves the Word document to file stream.
                    document.Save(outputFileStream, FormatType.Docx);
                }
            }
        }
    }
}
