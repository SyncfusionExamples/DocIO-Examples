using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;

namespace Add_continuous_sections_in_Word
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
                //Adds a paragraph to created section.
                IWParagraph paragraph = section.AddParagraph();
                string paraText = "AdventureWorks Cycles, the fictitious company on which the AdventureWorks sample databases are based, is a large, multinational manufacturing company.";
                //Appends the text to the created paragraph.
                paragraph.AppendText(paraText);
                //Adds the new section to the document.
                section = document.AddSection();
                //Sets a section break.
                section.BreakCode = SectionBreakCode.NoBreak;
                //Adds a paragraph to created section.
                paragraph = section.AddParagraph();
                //Appends the text to the created paragraph.
                paragraph.AppendText(paraText);
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
