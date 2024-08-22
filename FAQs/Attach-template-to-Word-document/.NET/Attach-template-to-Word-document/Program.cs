using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;

namespace Attach_template_to_Word_document
{
    class Program
    {
        static void Main(string[] args)
        {
            using (FileStream fileStreamPath = new FileStream(Path.GetFullPath(@"Data/Template.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                //Loads a source document.
                using (WordDocument document = new WordDocument(fileStreamPath, FormatType.Docx))
                {
                    //Attaches the template document to the source document.
                    document.AttachedTemplate.Path = @"D:/Data/Template.docx";
                    //Updates the styles of the document from the attached template each time the document is opened.
                    document.UpdateStylesOnOpen = true;
                    //Creates file stream.
                    using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"Output/Output.docx"), FileMode.Create, FileAccess.ReadWrite))
                    {
                        //Saves the Word document to file stream.
                        document.Save(outputFileStream, FormatType.Docx);
                    }
                }
            }
        }
    }
}
