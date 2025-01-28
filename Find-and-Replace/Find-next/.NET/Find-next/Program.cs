using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.Drawing;
using System.IO;

namespace Find_next
{
    class Program
    {
        static void Main(string[] args)
        {
            using (FileStream fileStreamPath = new FileStream(Path.GetFullPath(@"Data/Template.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                //Opens an existing Word document.
                using (WordDocument document = new WordDocument(fileStreamPath, FormatType.Automatic))
                {
                    //Finds the first occurrence of a particular text in the document.
                    TextSelection textSelection = document.Find("Adventure Works Cycles", false, true);
                    //Gets the found text as single text range.
                    WTextRange textRange = textSelection.GetAsOneRange();
                    //Modifies the text.
                    textRange.Text = "Replaced text";
                    //Sets highlight color.
                    textRange.CharacterFormat.HighlightColor = Color.Yellow;
                    //Finds the next occurrence of a particular text from the previous paragraph.
                    textSelection = document.FindNext(textRange.OwnerParagraph, "Adventure Works Cycles", true, false);
                    //Gets the found text as single text range.
                    WTextRange range = textSelection.GetAsOneRange();
                    //Sets bold formatting.
                    range.CharacterFormat.Bold = true;
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
}
