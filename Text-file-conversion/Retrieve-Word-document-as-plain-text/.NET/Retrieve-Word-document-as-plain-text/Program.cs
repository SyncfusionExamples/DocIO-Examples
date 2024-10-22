using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;

namespace Retrieve_Word_document_as_plain_text
{
    class Program
    {
        static void Main(string[] args)
        {
            //Loads an existing Word document into DocIO instance. 
            using (FileStream fileStreamPath = new FileStream(Path.GetFullPath(@"Data/Template.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                using (WordDocument document = new WordDocument(fileStreamPath, FormatType.Docx))
                {
                    //Gets the document text.
                    string text = document.GetText();
                    //Creates new Word document.
                    using (WordDocument newdocument = new WordDocument())
                    {
                        //Adds new section.
                        IWSection section = newdocument.AddSection();
                        //Adds new paragraph.
                        IWParagraph paragraph = section.AddParagraph();
                        //Appends the text to the paragraph.
                        paragraph.AppendText(text);
                        //Creates file stream.
                        using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"Output/Result.docx"), FileMode.Create, FileAccess.ReadWrite))
                        {
                            //Saves the Word document to file stream.
                            newdocument.Save(outputFileStream, FormatType.Docx);
                        }
                    }
                }
            }
        }
    }
}
