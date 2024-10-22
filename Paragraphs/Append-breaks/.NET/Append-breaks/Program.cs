using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;

namespace Append_breaks
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
                paragraph.AppendText("Before line break");
                //Adds line break to the paragraph.
                paragraph.AppendBreak(BreakType.LineBreak);
                paragraph.AppendText("After line break");
                IWParagraph pageBreakPara = section.AddParagraph();
                pageBreakPara.AppendText("Before page break");
                //Adds page break to the paragraph.
                pageBreakPara.AppendBreak(BreakType.PageBreak);
                pageBreakPara.AppendText("After page break");
                IWSection secondSection = document.AddSection();
                //Adds columns to the section.
                secondSection.AddColumn(100, 2);
                secondSection.AddColumn(100, 2);
                IWParagraph columnBreakPara = secondSection.AddParagraph();
                columnBreakPara.AppendText("Before column break");
                //Adds column break to the paragraph.
                columnBreakPara.AppendBreak(BreakType.ColumnBreak);
                columnBreakPara.AppendText("After column break");
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
