using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;

namespace Replace_text_of_a_text_range
{
    class Program
    {
        static void Main(string[] args)
        {
            using (FileStream fileStreamPath = new FileStream(Path.GetFullPath(@"../../../Template.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                //Creates a new Word document.
                using (WordDocument document = new WordDocument(fileStreamPath, FormatType.Automatic))
                {
                    //Gets the last paragraph.
                    WParagraph lastParagraph = document.LastParagraph;
                    //Iterates through the paragraph items to get the text range and modifies its content.
                    for (int i = 0; i < lastParagraph.ChildEntities.Count; i++)
                    {
                        if (lastParagraph.ChildEntities[i] is WTextRange)
                        {
                            WTextRange textRange = lastParagraph.ChildEntities[i] as WTextRange;
                            textRange.Text = "First text range of the last paragraph is replaced";
                            textRange.CharacterFormat.FontSize = 14;
                            break;
                        }
                    }
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
