using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.Office;
using System.IO;

namespace Set_Proofing_Language_to_Text
{
    class Program
    {
        static void Main(string[] args)
        {
            //Creates a new Word document.
            using (WordDocument document = new WordDocument())
            {
                //Add new section to the document.
                IWSection section = document.AddSection();
                //Add new paragraph to the section.
                IWParagraph paragraph = section.AddParagraph();
                //Add new text to the paragraph.
                IWTextRange text = paragraph.AppendText("Adventure Works Cycles, the fictitious company on which the AdventureWorks sample databases are based, is a large, multinational manufacturing company");
                //Set language identifier.
                text.CharacterFormat.LocaleIdASCII = (short)LocaleIDs.pt_PT;
                //Creates file stream.
                using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"../../../Output/Result.docx"), FileMode.Create, FileAccess.ReadWrite))
                {
                    //Saves the Word document to file stream.
                    document.Save(outputFileStream, FormatType.Docx);
                }
            }
        }
    }
}
