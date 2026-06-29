using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System;
using System.IO;

namespace Identify_Word_document_is_Encrypted_or_Protected
{
    class Program
    {
        static void Main(string[] args)
        {
            // Load a Word document.
            using (FileStream fileStream = new FileStream(Path.GetFullPath(@"Data/Template.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                try
                {
                    //Open the Word document using stream.
                    WordDocument document = new WordDocument(fileStream, FormatType.Docx);
                    ProtectionType protectionType = document.ProtectionType;

                    if (protectionType != ProtectionType.NoProtection)
                        Console.WriteLine("The Word document is protected by " + protectionType.ToString());
                    else
                        Console.WriteLine("The Word document is not protected.");
                }
                catch (Exception exception)
                {
                    //Return if Word document is encrypted.
                    if (exception.Message == "Document is encrypted, password is needed to open the document")
                    {
                        Console.WriteLine("The Word document is encrypted, need password to open.");
                    }
                }
            }
        }
    }
}