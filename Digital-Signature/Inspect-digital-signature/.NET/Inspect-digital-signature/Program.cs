using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.Office;

namespace Inspect_digital_signature
{
    class Program
    {
        static void Main(string[] args)
        {
            //Opens the signed Word document.
            WordDocument document = new WordDocument(Path.GetFullPath(@"Data\SignedDocument.docx"));
            // Gets the digital signature collection in the document.
            OfficeDigitalSignatureCollection signatures = document.DigitalSignatures;

            // Checks whether every digital signature in the collection is valid.
            bool allValid = signatures.IsValid;
            Console.WriteLine("All signatures are valid: " + allValid);

            // Displays details of each signature.
            foreach (OfficeDigitalSignature signature in signatures)
            {
                Console.WriteLine("Signature comments: " + signature.Comments);
                Console.WriteLine("Time of signing: " + signature.SigningTime);
                Console.WriteLine("Subject name: " +
                    signature.Certificate.Subject);
                Console.WriteLine("Issuer name: " +
                    signature.Certificate.Issuer);
                Console.WriteLine();
            }

            // Closes the document.
            document.Close();
        }
    }
}

