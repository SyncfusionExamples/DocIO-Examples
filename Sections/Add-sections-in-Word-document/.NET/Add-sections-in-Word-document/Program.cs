using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;

namespace Add_sections_in_Word_document
{
    class Program
    {
        static void Main(string[] args)
        {
            //Creates a new instance of WordDocument (Empty Word Document).
            using (WordDocument document = new WordDocument())
            {
                //Adds new section to the document.
                IWSection section = document.AddSection();
                //Adds new paragraph to the section.
                IWParagraph paragraph = section.AddParagraph();
                //Appends the text to the created paragraph.
                paragraph.AppendText("AdventureWorks Cycles, the fictitious company on which the AdventureWorks sample databases are based, is a large, multinational manufacturing company.");
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
