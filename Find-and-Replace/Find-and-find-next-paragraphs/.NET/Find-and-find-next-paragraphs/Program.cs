using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.Drawing;
using System.IO;

namespace Find_and_find_next_paragraphs
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
                    //Finds the first occurrence of a particular text extended to several paragraphs in the document.
                    TextSelection[] textSelections = document.FindSingleLine("First paragraph Second paragraph", true, false);
                    WParagraph paragraph = null;
                    foreach (TextSelection textSelection in textSelections)
                    {
                        //Gets the found text as single text range and set highlight color.
                        WTextRange textRange = textSelection.GetAsOneRange();
                        textRange.CharacterFormat.HighlightColor = Color.YellowGreen;
                        paragraph = textRange.OwnerParagraph;
                    }
                    //Finds the next occurrence of a particular text extended to several paragraphs in the document.
                    textSelections = document.FindNextSingleLine(paragraph, "First paragraph Second paragraph", true, false);
                    foreach (TextSelection textSelection in textSelections)
                    {
                        //Gets the found text as single text range and sets italic formatting.
                        WTextRange text = textSelection.GetAsOneRange();
                        text.CharacterFormat.Italic = true;
                    }
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
