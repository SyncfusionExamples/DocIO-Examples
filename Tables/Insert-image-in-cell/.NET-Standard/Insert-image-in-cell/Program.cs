using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;

namespace Insert_image_in_cell
{
    class Program
    {
        static void Main(string[] args)
        {
            //Opens an existing Word document.
            using (WordDocument document = new WordDocument())
            {
                IWSection section = document.AddSection();
                //Adds a new table into Word document.
                IWTable table = section.AddTable();
                //Specifies the total number of rows & columns.
                table.ResetCells(2, 2);
                table[0, 0].AddParagraph().AppendText("Product Name");
                table[0, 1].AddParagraph().AppendText("Product Image");
                table[1, 0].AddParagraph().AppendText("Apple Juice");
                //Adds the image into cell.
                FileStream imageStream = new FileStream(Path.GetFullPath(@"../../../Image.png"), FileMode.Open, FileAccess.ReadWrite);
                IWPicture picture = table[1, 1].AddParagraph().AppendPicture(imageStream);
                picture.Height = 75;
                picture.Width = 60;
                //Creates file stream.
                using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"../../../Result.docx"), FileMode.Create, FileAccess.ReadWrite))
                {
                    //Saves the Word document to file stream.
                    document.Save(outputFileStream, FormatType.Docx);
                }
            }
        }
    }
}
