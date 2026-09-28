using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.Office;

namespace Validate_digital_signature
{
    class Program
    {
        static void Main(string[] args)
        {
            //Opens the signed Word document.
            WordDocument document = new WordDocument(Path.GetFullPath(@"Data\SignedDocument.docx"));
            //Gets the digital signature collection in the document.
            OfficeDigitalSignatureCollection signatures = document.DigitalSignatures;
            //Checks whether every digital signature in the collection is valid.
            bool allValid = signatures.IsValid;
            Console.WriteLine("All signatures are valid : " + allValid);
            //Iterates through each signature and checks whether it is valid.
            foreach (OfficeDigitalSignature signature in signatures)
            {
                //Checks whether the signature is valid.
                bool isValid = signature.IsValid;
                Console.WriteLine("Signature is valid : " + isValid);
            }
            //Closes the document
            document.Close();

        }
    }
}

