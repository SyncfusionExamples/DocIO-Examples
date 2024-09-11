using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;
using System.Text.RegularExpressions;

namespace Find_and_replace_text_with_formatted_text
{
    class Program
    {
        static void Main(string[] args)
        {
            using (FileStream fileStream = new FileStream(Path.GetFullPath(@"Data/Template.docx"), FileMode.Open, FileAccess.ReadWrite))
            {
                //Open an existing Word document.
                using (WordDocument document = new WordDocument(fileStream, FormatType.Automatic))
                {
                    //Find the first occurrence of a particular text in the document.
                    TextSelection selection = document.Find(new Regex ("^«(.*)»"));
                    //Replace the particular text with the selected text along with formatting.
                    document.Replace("Bear", selection, false, false, true);
                    //Create file stream.
                    using (FileStream outputStream = new FileStream(Path.GetFullPath(@"Output/Result.docx"), FileMode.Create, FileAccess.ReadWrite))
                    {
                        //Save the Word document to file stream.
                        document.Save(outputStream, FormatType.Docx);
                    }
                }
            }
        }
    }
}
