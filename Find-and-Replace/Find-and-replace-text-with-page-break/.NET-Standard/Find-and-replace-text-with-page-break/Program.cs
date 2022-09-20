using System.IO;
using System.Text.RegularExpressions;
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;

namespace Find_and_replace_text_with_page_break
{
    class Program
    {
        static void Main(string[] args)
        {
            //Create file stream.
            using (FileStream fileStreamPath = new FileStream(Path.GetFullPath(@"../../../Data/Input.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                //Open an existing Word document.
                using (WordDocument document = new WordDocument(fileStreamPath, FormatType.Docx))
                {
                    //Add new paragraph to the section.
                    WParagraph paragraph = document.Sections[0].AddParagraph() as WParagraph;
                    //Add the page break.
                    paragraph.AppendBreak(BreakType.PageBreak);
                    //Create text body part.
                    TextBodyPart bodyPart = new TextBodyPart(document);
                    bodyPart.BodyItems.Add(paragraph);
                    //Replace all entries of a given regular expression text with the text body part.
                    document.ReplaceSingleLine(new Regex("<<(.*)>>"), bodyPart);
                    //Creates file stream.
                    using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"../../../Sample.docx"), FileMode.Create, FileAccess.ReadWrite))
                    {
                        //Save the Word document to file stream.
                        document.Save(outputFileStream, FormatType.Docx);
                    }
                }
            }
        }
    }
}
