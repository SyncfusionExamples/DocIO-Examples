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

            // Gets the number of digital signatures in the document.
            int signatureCount = signatures.Count;
            Console.WriteLine("Signature count: " + signatureCount);

            // Checks whether every digital signature in the collection is valid.
            bool allValid = signatures.IsValid;
            Console.WriteLine("All signatures are valid: " + allValid);

            // Displays details of each signature.
            foreach (OfficeDigitalSignature signature in signatures)
            {
                // Checks whether the signature is valid.
                bool isValid = signature.IsValid;
                // Gets the comments recorded with the signature.
                string comments = signature.Comments;
                // Gets the timestamp recorded with the signature.
                DateTime signingTime = signature.SigningTime;
                // Gets the application version recorded with the signature.
                string applicationVersion = signature.ApplicationVersion;
                // Gets the Office version recorded with the signature.
                string officeVersion = signature.OfficeVersion;
                // Gets the Windows version recorded with the signature.
                string windowsVersion = signature.WindowsVersion;
                // Gets the color depth recorded with the signature.
                int colorDepth = signature.ColorDepth;
                // Gets the horizontal resolution recorded with the signature.
                float horizontalResolution = signature.HorizontalResolution;
                // Gets the vertical resolution recorded with the signature.
                float verticalResolution = signature.VerticalResolution;
                // Gets the raw signature value bytes.
                byte[] signatureValue = signature.SignatureValue;
                // Gets the certificate used to create the signature.
                OfficeDigitalSignatureCertificate certificate = signature.Certificate;

                Console.WriteLine("Signature is valid: " + isValid);
                Console.WriteLine("Comments: " + comments);
                Console.WriteLine("Signing time: " + signingTime);
                Console.WriteLine("Application version: " + applicationVersion);
                Console.WriteLine("Office version: " + officeVersion);
                Console.WriteLine("Windows version: " + windowsVersion);
                Console.WriteLine("Color depth: " + colorDepth);
                Console.WriteLine("Horizontal resolution: " + horizontalResolution);
                Console.WriteLine("Vertical resolution: " + verticalResolution);
                Console.WriteLine("Signature value length: " + (signatureValue != null ? signatureValue.Length : 0));

                if (certificate != null)
                {
                    // Gets the certificate subject.
                    string subject = certificate.Subject;
                    // Gets the certificate issuer.
                    string issuer = certificate.Issuer;
                    // Gets the certificate serial number.
                    string serialNumber = certificate.SerialNumber;
                    // Gets the certificate thumbprint.
                    string thumbprint = certificate.Thumbprint;
                    // Gets the date the certificate is valid from.
                    DateTime validFrom = certificate.ValidFrom;
                    // Gets the date the certificate is valid to.
                    DateTime validTo = certificate.ValidTo;

                    Console.WriteLine("Subject name: " + subject);
                    Console.WriteLine("Issuer name: " + issuer);
                    Console.WriteLine("Serial number: " + serialNumber);
                    Console.WriteLine("Thumbprint: " + thumbprint);
                    Console.WriteLine("Valid from: " + validFrom);
                    Console.WriteLine("Valid to: " + validTo);
                }

                Console.WriteLine();
            }

            // Retrieves a specific signature by index from the collection.
            if (signatures.Count > 0)
            {
                OfficeDigitalSignature firstSignature = signatures[0];
                Console.WriteLine("First signature comments: " + firstSignature.Comments);
            }

            // Walks the document body and reads the properties of any signature line shapes.
            foreach (WSection section in document.Sections)
            {
                foreach (WParagraph paragraph in section.Paragraphs)
                {
                    foreach (Entity entity in paragraph.ChildEntities)
                    {
                        if (entity is WPicture picture && picture.IsSignatureLine)
                        {
                            OfficeSignatureLine signatureLine = picture.SignatureLine;

                            // Gets the unique identifier of the signature line.
                            Guid id = signatureLine.Id;
                            // Gets the instructions displayed to the signer.
                            string instructions = signatureLine.Instructions;
                            // Gets whether the signer can attach comments when signing.
                            bool allowComments = signatureLine.AllowComments;
                            // Gets whether the sign date is displayed on the signature line.
                            bool showDate = signatureLine.ShowDate;
                            // Gets whether the signature line has been signed.
                            bool isSigned = signatureLine.IsSigned;

                            Console.WriteLine("Signature line id: " + id);
                            Console.WriteLine("Signature line instructions: " + instructions);
                            Console.WriteLine("Signature line allow comments: " + allowComments);
                            Console.WriteLine("Signature line show date: " + showDate);
                            Console.WriteLine("Signature line is signed: " + isSigned);
                        }
                    }
                }
            }

            // Closes the document.
            document.Close();
        }
    }
}

