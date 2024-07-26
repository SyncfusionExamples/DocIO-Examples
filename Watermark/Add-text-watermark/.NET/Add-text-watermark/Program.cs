using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.Drawing;
using System.IO;

namespace Add_text_watermark
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
                paragraph.AppendText("AdventureWorks Cycles, the fictitious company on which the AdventureWorks sample databases are based, is a large, multinational manufacturing company.");
                //Creates a new text watermark.
                TextWatermark textWatermark = new TextWatermark("TextWatermark", "", 250, 100);
                //Sets the created watermark to the document.
                document.Watermark = textWatermark;
                //Sets the text watermark font size.
                textWatermark.Size = 72;
                //Sets the text watermark layout to Horizontal.
                textWatermark.Layout = WatermarkLayout.Horizontal;
                textWatermark.Semitransparent = false;
                //Sets the text watermark text color.
                textWatermark.Color = Color.Black;
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
