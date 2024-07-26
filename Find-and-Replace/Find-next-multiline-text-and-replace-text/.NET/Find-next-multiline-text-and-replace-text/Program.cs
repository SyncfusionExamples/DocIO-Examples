using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;
using System.Text.RegularExpressions;

namespace Find_next_multiline_text_and_replace_text
{
    class Program
    {
        static void Main(string[] args)
        {
            using (FileStream fileStream = new FileStream(Path.GetFullPath(@"../../../Template.docx"), FileMode.Open, FileAccess.ReadWrite))
            {
                //Load the file stream into a Word document.
                using (WordDocument document = new WordDocument(fileStream, FormatType.Automatic))
                {
                    //Access the specific table in a Word document.
                    WTable table = document.LastSection.Tables[0] as WTable;
                    //Find the next occurrence of particular text extended to several paragraphs after the specific table.
                    TextSelection[] textSelections = document.FindNextSingleLine(table, new Regex(@"\[(.*)\]"));
                    //Replace the particular text with the selected text.
                    document.Replace("Equation of sodium chloride and silver nitrate", textSelections[1], true, true);
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
