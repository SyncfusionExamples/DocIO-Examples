using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;

namespace Find_and_modify_hyperlink_address
{
    class Program
    {
        static void Main(string[] args)
        {
            //Open the file as a stream.
            using (FileStream inputStream = new FileStream(Path.GetFullPath(@"../../../Input.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                //Load the file stream into a Word document.
                using (WordDocument document = new WordDocument(inputStream, FormatType.Docx))
                {
                    //Access paragraph in a Word document.
                    WParagraph paragraph = document.Sections[0].Paragraphs[1];
                    WField field = paragraph.ChildEntities[0] as WField;
                    //Create an instance of hyperlink.
                    Hyperlink hyperlink = new Hyperlink(field);
                    //Set the hyperlink type, URL and the text to display.
                    hyperlink.Type = HyperlinkType.WebLink;
                    hyperlink.Uri = "http://www.google.com";
                    hyperlink.TextToDisplay = "Google";
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
