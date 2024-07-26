using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;

namespace Code_block_in_Markdown
{
    class Program
    {
        static void Main(string[] args)
        {
            //Create a new Word document.
            using (WordDocument document = new WordDocument())
            {
                //Add a new section to the document.
                IWSection section = document.AddSection();
                //Add a new paragraph to the section.
                IWParagraph paragraph = section.AddParagraph();
                //Append text to the paragraph.
                IWTextRange textRange = paragraph.AppendText("Fenced Code");
                //Add a new paragraph to the section.
                paragraph = section.AddParagraph();
                //Create a user-defined style as FencedCode.
                IWParagraphStyle style = document.AddParagraphStyle("FencedCode");
                //Apply FencedCode style for the paragraph.
                paragraph.ApplyStyle("FencedCode");
                //Append text.
                textRange = paragraph.AppendText("class Hello\n{\n\tStatic void Main()\n\t{\n\t\tConsole.WriteLine(\"Fenced Code\")\n\t}\n}");
                //Add a new paragraph and append text to the paragraph.
                section.AddParagraph().AppendText("Indented Code");
                //Add a new paragraph to the section.
                paragraph = section.AddParagraph();
                //Create a user-defined style as IndentedCode.
                style = document.AddParagraphStyle("IndentedCode");
                //Apply IndentedCode style for the paragraph.
                paragraph.ApplyStyle("IndentedCode");
                //Append text.
                textRange = paragraph.AppendText("class Hello\n\t{\n\t\tStatic void Main()\n\t\t{\n\t\t\tConsole.WriteLine(\"Indented Code\")\n\t\t}\n\t}");
                //Create a file stream.
                using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"../../../WordToMd.md"), FileMode.Create, FileAccess.ReadWrite))
                {
                    //Save the Markdown file to the file stream.
                    document.Save(outputFileStream, FormatType.Markdown);
                }
            }
        }
    }
}
