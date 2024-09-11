using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;

namespace Increase_or_decrease_list_indent
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
                //Applies default numbered list style.
                paragraph.ListFormat.ApplyDefNumberedStyle();
                //Adds text to the paragraph.
                paragraph.AppendText("Multilevel numbered list - Level 0");
                //Continues the list defined.
                paragraph.ListFormat.ContinueListNumbering();
                //Adds second paragraph.
                paragraph = section.AddParagraph();
                paragraph.AppendText("Multilevel numbered list - Level 1");
                //Continues last defined list.
                paragraph.ListFormat.ContinueListNumbering();
                //Increases the level indent.
                paragraph.ListFormat.IncreaseIndentLevel();
                //Adds new paragraph.
                paragraph = section.AddParagraph();
                paragraph.AppendText("Multilevel numbered list - Level 0");
                //Continues last defined list.
                paragraph.ListFormat.ContinueListNumbering();
                //Increases the level indent.
                paragraph.ListFormat.DecreaseIndentLevel();
                //Adds new paragraph.
                paragraph = section.AddParagraph();
                paragraph.AppendText("Multilevel numbered list - Level 1");
                //Continues last defined list.
                paragraph.ListFormat.ContinueListNumbering();
                //Increases the level indent.
                paragraph.ListFormat.IncreaseIndentLevel();
                //Creates file stream.
                using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"Output/Result.docx"), FileMode.Create, FileAccess.ReadWrite))
                {
                    //Saves the Word document to file stream.
                    document.Save(outputFileStream, FormatType.Docx);
                }
            }
        }
    }
}
