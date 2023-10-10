using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;

namespace Add_svg_image
{
    class Program
    {
        static void Main(string[] args)
        {
            //Create a new Word document.
            using (WordDocument document = new WordDocument())
            {                   
                     //Add new section to the document.
                     IWSection section = document.AddSection();
                     //Add new paragraph to the section.
                     IWParagraph firstParagraph = section.AddParagraph();
                     //Get the image as byte array.
                     byte[] imageBytes = File.ReadAllBytes(Path.GetFullPath(@"../../../Data/Buyers.png"));
                     //Get the SVG image as byte array.
                     byte[] svgData = File.ReadAllBytes(Path.GetFullPath(@"../../../Data/Buyers.svg"));
                     //Add SVG image to the paragraph.
                     IWPicture picture = firstParagraph.AppendPicture(svgData, imageBytes);
                     //Set height and width for the image.
                     picture.Height = 100;
                     picture.Width = 100;

                //Create a file stream.
                using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"../../../Result.docx"), FileMode.Create, FileAccess.ReadWrite))
                {
                    //Save the Word document to file stream.
                    document.Save(outputFileStream, FormatType.Docx);
                }
            }
        }
    }
}
