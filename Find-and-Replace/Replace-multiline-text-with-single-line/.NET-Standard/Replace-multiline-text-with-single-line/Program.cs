using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;
using System.Text.RegularExpressions;

namespace Replace_multiline_text_with_single_line
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
                    //Replace the text extended to several paragraphs with simple text.
                    document.ReplaceSingleLine(new Regex("«(.*)»"), "Replaced paragraph");
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
