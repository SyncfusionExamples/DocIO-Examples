using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;

namespace Simple_bulleted_list
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
                paragraph.ListFormat.ApplyDefBulletStyle();
                //Adds text to the paragraph.
                paragraph.AppendText("List item 1");
                //Continues the list defined.
                paragraph.ListFormat.ContinueListNumbering();
                //Adds second paragraph.
                paragraph = section.AddParagraph();
                paragraph.AppendText("List item 2");
                //Continues last defined list.
                paragraph.ListFormat.ContinueListNumbering();
                //Adds new paragraph.
                paragraph = section.AddParagraph();
                paragraph.AppendText("List item 3");
                //Continues last defined list.
                paragraph.ListFormat.ContinueListNumbering();
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
