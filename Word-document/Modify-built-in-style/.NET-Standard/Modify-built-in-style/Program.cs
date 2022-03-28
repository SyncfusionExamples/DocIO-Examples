using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.Drawing;
using System.IO;

namespace Modify_built_in_style
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
                //Creates built-in style and modifies its properties.
                Style style = Style.CreateBuiltinStyle(BuiltinStyle.Heading1, document) as Style;
                style.CharacterFormat.Italic = true;
                style.CharacterFormat.TextColor = Color.DarkGreen;
                //Adds it to the styles collection.
                document.Styles.Add(style);
                //Adds new paragraph to the section.
                IWParagraph paragraph = section.AddParagraph();
                IWTextRange text = paragraph.AppendText("A built-in style is modified and is applied to this paragraph.");
                //Applies the new style to paragraph.
                paragraph.ApplyStyle(style.Name);
                //Creates file stream.s
                using (FileStream outputStream = new FileStream(Path.GetFullPath(@"../../../Result.docx"), FileMode.Create, FileAccess.ReadWrite, FileShare.ReadWrite))
                {
                    //Saves the Word document to file stream.s
                    document.Save(outputStream, FormatType.Docx);
                }
            }
        }
    }
}
