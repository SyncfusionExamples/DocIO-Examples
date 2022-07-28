using System.IO;
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;

namespace Text_wrapping_break
{
    class Program
    {
        static void Main(string[] args)
        {
            using (FileStream fileStream = new FileStream(Path.GetFullPath(@"../../../Template.docx"), FileMode.Open))
            {
                //Loads an existing Word document.
                using (WordDocument document = new WordDocument(fileStream, FormatType.Automatic))
                {
                    //Access paragraph from section.
                    WParagraph paragraph = document.LastSection.Body.ChildEntities[2] as WParagraph;
                    //Create text wrapping break.
                    Break textWrappingBreak = new Break(document, BreakType.TextWrappingBreak);
                    //Insert text wrapping break in specific index.
                    paragraph.ChildEntities.Insert(1, textWrappingBreak);
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
}
