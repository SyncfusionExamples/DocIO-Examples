using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System;
using System.IO;

namespace Identify_Word_document_is_Encrypted_or_not
{
    class Program
    {
        static void Main(string[] args)
        {
            //Creates a new Word document.
            using (FileStream fileStreamPath = new FileStream(Path.GetFullPath(@"Data/Template.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                Console.WriteLine(IsEncrypted(fileStreamPath));
            }
        }
        public static string IsEncrypted(FileStream fileStream)
        {
            try
            {
                //Open the existing Word document.
                WordDocument document = new WordDocument(fileStream, FormatType.Docx);
                return "Document is not encrypted.";
            }
            catch (Exception exception)
            {
                //Return if Word document is encrypted.
                if (exception.Message == "Document is encrypted, password is needed to open the document")
                {
                    return exception.Message;
                }
            }
            return "Document is not encrypted.";
        }
    }
}