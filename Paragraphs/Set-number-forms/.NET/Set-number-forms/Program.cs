using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;

namespace Set_number_forms
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
                //Adds new text.
                IWTextRange text = paragraph.AppendText("Numbers to describe oldstyle number form 0123456789");
                text.CharacterFormat.FontName = "Calibri";
                //Sets number style.
                text.CharacterFormat.NumberForm = NumberFormType.OldStyle;
                paragraph = section.AddParagraph();
                text = paragraph.AppendText("Numbers to describe lining number form 0123456789");
                text.CharacterFormat.FontName = "Calibri";
                //Sets number style.
                text.CharacterFormat.NumberForm = NumberFormType.Lining;
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
