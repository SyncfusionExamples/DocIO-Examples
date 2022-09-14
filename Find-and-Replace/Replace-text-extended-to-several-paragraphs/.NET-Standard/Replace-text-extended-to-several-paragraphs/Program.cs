using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;
using System.Text.RegularExpressions;

namespace Replace_text_extended_to_several_paragraphs
{
    class Program
    {
        static void Main(string[] args)
        {
            using (FileStream fileStream = new FileStream(Path.GetFullPath(@"../../../Template.docx"), FileMode.Open, FileAccess.ReadWrite))
            {
                //Open an existing Word document.
                using (WordDocument document = new WordDocument(fileStream, FormatType.Automatic))
                {
                    //Find the first occurrence of particular text extended to several paragraphs in the document.
                    TextSelection[] textSelections = document.FindSingleLine(new Regex ("«(.*)»"));
                    //Replace the particular text extended to several paragraphs with the selected text.
                    document.ReplaceSingleLine(new Regex(@"\[(.*)\]"), textSelections[1]);
                    //Create file stream.
                    using (FileStream outputStream = new FileStream(Path.GetFullPath(@"../../../Result.docx"), FileMode.Create, FileAccess.ReadWrite))
                    {
                        //Save the Word document to file stream.
                        document.Save(outputStream, FormatType.Docx);
                    }
                }
            }
        }
    }
}
