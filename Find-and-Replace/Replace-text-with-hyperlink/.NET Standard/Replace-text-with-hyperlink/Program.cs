using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System;
using System.IO;

namespace Replace_text_with_hyperlink
{
    internal class Program
    {
        static void Main(string[] args)
        {
            using (FileStream fileStreamPath = new FileStream(Path.GetFullPath(@"../../../Data/Input.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                //Opens an existing Word document.
                using (WordDocument document = new WordDocument(fileStreamPath, FormatType.Docx))
                {
                    string textToReplace = "Syncfusion";
                    //Initialize new text body part
                    TextBodyPart textBodyPart = new TextBodyPart(document);
                    //Initialize new paragraph
                    WParagraph paragraph = new WParagraph(document);
                    //Adds the paragraph into the text body part
                    textBodyPart.BodyItems.Add(paragraph);
                    //Appends web hyperlink to the paragraph
                    paragraph.AppendHyperlink("http://www.syncfusion.com", textToReplace, HyperlinkType.WebLink);
                    //Replaces all entries of given string in the document with text body part
                    document.Replace(textToReplace, textBodyPart, false, true);
                    //Clears the text body part
                    textBodyPart.Clear();
                    //Creates file stream.
                    using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"../../../Data/Result.docx"), FileMode.Create, FileAccess.ReadWrite))
                    {
                        //Saves the Word document to file stream.
                        document.Save(outputFileStream, FormatType.Docx);
                    }
                }
            }
        }
    }
}
