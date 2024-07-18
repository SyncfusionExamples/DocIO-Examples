using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.Drawing;
using System.IO;

namespace Create_custom_paragraph_style
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
                //Creates user defined style.
                IWParagraphStyle style = document.AddParagraphStyle("User_defined_style");
                style.ParagraphFormat.BackColor = Color.LightGray;
                style.ParagraphFormat.AfterSpacing = 18f;
                style.ParagraphFormat.BeforeSpacing = 18f;
                style.ParagraphFormat.Borders.BorderType = Syncfusion.DocIO.DLS.BorderStyle.DotDash;
                style.ParagraphFormat.Borders.LineWidth = 0.5f;
                style.ParagraphFormat.LineSpacing = 15f;
                style.CharacterFormat.FontName = "Calibri";
                style.CharacterFormat.Italic = true;
                //Adds new paragraph to the section.
                IWParagraph paragraph = section.AddParagraph();
                IWTextRange text = paragraph.AppendText("A new paragraph style is created and is applied to this paragraph.");
                //Applies the new style to paragraph.
                paragraph.ApplyStyle("User_defined_style");
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
