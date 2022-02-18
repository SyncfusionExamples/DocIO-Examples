using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.Drawing;
using System.IO;

namespace Apply_paragraph_formatting
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
                //Adds new paragraph to the section.
                IWParagraph paragraph = section.AddParagraph();
                //Adds new text to the paragraph.
                IWTextRange firstText = paragraph.AppendText("Paragraphs are the basic elements of the Word document. It can contain text and images.");
                //Applies paragraph formatting.
                paragraph.ParagraphFormat.AfterSpacing = 18f;
                paragraph.ParagraphFormat.BeforeSpacing = 18f;
                paragraph.ParagraphFormat.BackColor = Color.LightGray;
                paragraph.ParagraphFormat.FirstLineIndent = 10f;
                paragraph.ParagraphFormat.LineSpacing = 10f;
                paragraph.ParagraphFormat.HorizontalAlignment = Syncfusion.DocIO.DLS.HorizontalAlignment.Right;
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
