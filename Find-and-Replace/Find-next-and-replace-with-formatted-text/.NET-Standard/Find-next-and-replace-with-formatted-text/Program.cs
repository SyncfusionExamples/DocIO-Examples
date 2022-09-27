using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;
using System.Text.RegularExpressions;

namespace Find_next_and_replace_with_formatted_text
{
    class Program
    {
        static void Main(string[] args)
        {
            //Open the file as stream.
            using (FileStream inputStream = new FileStream(Path.GetFullPath(@"../../../Input.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                //Load the file stream into a Word document.
                using (WordDocument document = new WordDocument(inputStream, FormatType.Docx))
                {
                    //Access the paragraph in the Word document.
                    TextBodyItem textBodyItem = document.Sections[0].Paragraphs[3] as WParagraph;
                    //Get the next entry of the specified regex from the text body item.
                    TextSelection textSelections = document.FindNext(textBodyItem, new Regex("Adventure Works Cycles"));
                    //Find the text that extends to several paragraphs and replace it with the desired content.
                    document.ReplaceSingleLine("CompanyName", textSelections, true, true);
                    //Create a file stream.
                    using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"../../../Sample.docx"), FileMode.Create, FileAccess.ReadWrite))
                    {
                        //Save the Word document to the file stream.
                        document.Save(outputFileStream, FormatType.Docx);
                    }
                }
            }
        }
    }
}
