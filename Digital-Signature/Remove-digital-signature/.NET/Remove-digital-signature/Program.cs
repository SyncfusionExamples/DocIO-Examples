using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.Office;

namespace Remove_digital_signature
{
    class Program
    {
        static void Main(string[] args)
        {
            //Opens the signed Word document.
            WordDocument document = new WordDocument(Path.GetFullPath(@"Data\SignedDocument.docx"));
            //Removes all digital signatures from the document.
            document.RemoveAllDigitalSignatures();
            //Saves the Word document to file.
            document.Save(Path.GetFullPath(@"Output\Result.docx"), FormatType.Docx);
            //Closes the document
            document.Close();
        }
    }
}

