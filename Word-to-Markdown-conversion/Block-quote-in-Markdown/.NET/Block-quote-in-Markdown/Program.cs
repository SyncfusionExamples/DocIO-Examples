using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;

namespace Block_quote_in_Markdown
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
                //Create a user-defined style.
                IWParagraphStyle style = document.AddParagraphStyle("Quote");
                //Add a new paragraph to the section.
                IWParagraph paragraph = section.AddParagraph();
                //Apply Quote style to simple hello world text.
                paragraph.ApplyStyle("Quote");
                //Append text.
                IWTextRange textRange = paragraph.AppendText("Hello World");
                //Save the document as a Markdown file.
                document.Save(Path.GetFullPath(@"Output/Output.md"));
            }
        }
    }
}
