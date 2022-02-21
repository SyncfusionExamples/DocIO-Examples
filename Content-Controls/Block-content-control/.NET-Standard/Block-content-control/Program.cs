using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;

namespace Block_content_control
{
    class Program
    {
        static void Main(string[] args)
        {
            //Creates a new Word document.
            using (WordDocument document = new WordDocument())
            {
                //Adds new section to the document.
                IWSection section = document.AddSection();
                WTextBody textBody = section.Body;
                //Adds block content control into Word document.
                BlockContentControl blockContentControl = textBody.AddBlockContentControl(ContentControlType.RichText) as BlockContentControl;
                //Adds new paragraph in the block content control.
                WParagraph paragraph = blockContentControl.TextBody.AddParagraph() as WParagraph;
                //Adds new text to the paragraph.
                paragraph.AppendText("Block content control");
                //Adds new table to the block content control.
                WTable table = blockContentControl.TextBody.AddTable() as WTable;
                //Specifies the total number of rows and columns.
                table.ResetCells(2, 3);
                //Adds new paragraph to the block content control.
                paragraph = blockContentControl.TextBody.AddParagraph() as WParagraph;
                //Gets the image stream.
                FileStream imageStream = new FileStream(Path.GetFullPath(@"../../../Image.png"), FileMode.Open, FileAccess.Read);
                //Adds image to the paragraph.
                paragraph.AppendPicture(imageStream);
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
