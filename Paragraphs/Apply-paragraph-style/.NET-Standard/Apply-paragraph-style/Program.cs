using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.Drawing;
using System.IO;

namespace Apply_paragraph_style
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
                IWParagraph firstParagraph = section.AddParagraph();
                //Adds new text to the paragraph.
                IWTextRange firstText = firstParagraph.AppendText("Built-in styles can be applied to the paragraph. Heading1 style is applied to this paragraph.");
                //Applies built-in style for the paragraph.
                firstParagraph.ApplyStyle(BuiltinStyle.Heading1);
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
