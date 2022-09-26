using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.Drawing;
using System;
using System.IO;
using System.Text.RegularExpressions;
using System.Xml;

namespace Find_Next_and_replace_with_formatted_text
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
                    //Get the text body item.
                    TextBodyItem textBodyItem = document.Sections[0].Paragraphs[3] as WParagraph;
                    //Get the next entry of the specifoed regex from the text body item.
                    TextSelection textSelections = document.FindNext(textBodyItem, new Regex("Adventure Works Cycles"));
                    //Get the found text as single text range and format it.
                    WTextRange textRange = textSelections.GetAsOneRange();
                    textRange.CharacterFormat.Bold = true;
                    textRange.CharacterFormat.FontName = "Times New Roman";
                    textRange.CharacterFormat.FontSize = 12;
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
