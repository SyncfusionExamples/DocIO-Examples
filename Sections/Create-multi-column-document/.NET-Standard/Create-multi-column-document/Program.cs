using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;

namespace Create_multi_column_document
{
    class Program
    {
        static void Main(string[] args)
        {
            //Creates a new Word document.
            using (WordDocument document = new WordDocument())
            {
                //Adds the section into Word document.
                IWSection section = document.AddSection();
                //Adds the column into the section.
                section.AddColumn(150, 20);
                //Adds the column into the section.
                section.AddColumn(150, 20);
                //Adds the column into the section.
                section.AddColumn(150, 20);
                //Adds a paragraph to created section.
                IWParagraph paragraph = section.AddParagraph();
                //Adds a paragraph to created section.
                paragraph = section.AddParagraph();
                string paraText = "AdventureWorks Cycles, the fictitious company on which the AdventureWorks sample databases are based, is a large, multinational manufacturing company.";
                //Appends the text to the created paragraph.
                paragraph.AppendText(paraText);
                //Adds the column break.
                paragraph.AppendBreak(BreakType.ColumnBreak);
                //Adds a paragraph to created section.
                paragraph = section.AddParagraph();
                //Appends the text to the created paragraph.
                paragraph.AppendText(paraText);
                //Adds the column break.
                paragraph.AppendBreak(BreakType.ColumnBreak);
                //Adds a paragraph to created section.
                paragraph = section.AddParagraph();
                //Appends the text to the created paragraph.
                paragraph.AppendText(paraText);
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
