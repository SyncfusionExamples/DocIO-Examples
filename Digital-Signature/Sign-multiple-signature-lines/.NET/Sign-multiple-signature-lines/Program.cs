using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.Office;

namespace Sign_multiple_signature_lines
{
    class Program
    {
        static void Main(string[] args)
        {
            // Opens an existing Word document and adds signature lines.
            Dictionary<Guid, string> signatureInfo;
            using (WordDocument document = new WordDocument(Path.GetFullPath(@"Data\Template.docx")))
            {
                // Gets the last section of the document.
                IWSection section = document.LastSection;

                // Stores the signature line IDs and corresponding signature image paths.
                signatureInfo = new Dictionary<Guid, string>();

                // Defines the signers.
                string[] signers = { "Tony", "Steve", "Bruce" };

                // Defines the signature images corresponding to each signer.
                string[] images =
                {
                    Path.GetFullPath(@"Data\TonySignature.png"),
                    Path.GetFullPath(@"Data\SteveSignature.png"),
                    Path.GetFullPath(@"Data\BruceSignature.png")
                };

                // Adds a signature line for each signer.
                for (int i = 0; i < signers.Length; i++)
                {
                    IWParagraph paragraph = section.AddParagraph();

                    IWPicture picture = paragraph.AppendSignatureLine(
                        new SignatureLineSettings()
                        {
                            Signer = signers[i],
                            SignerTitle = "Approver",
                            Email = signers[i] + "@example.com",
                            Instructions = "Please review and sign.",
                            ShowDate = true
                        },
                        192,
                        96);

                    // Gets the unique identifier of the inserted signature line.
                    Guid signatureLineId = ((WPicture)picture).SignatureLine.Id;

                    // Maps the signature line to its corresponding signature image.
                    signatureInfo.Add(signatureLineId, images[i]);
                }

                // Saves the document after adding the signature lines.
                document.Save(Path.GetFullPath(@"Output\Result.docx"), FormatType.Docx);
            }

            // Loads the signing certificate.
            OfficeDigitalSignatureCertificate certificate =
                new OfficeDigitalSignatureCertificate(
                    Path.GetFullPath(@"Data\certificate.pfx"),
                    "password123");

            // Opens the saved document to sign each signature line.
            using (WordDocument document = new WordDocument(Path.GetFullPath(@"Output\Result.docx")))
            {
                // Signs each signature line using its corresponding signature image.
                foreach (KeyValuePair<Guid, string> item in signatureInfo)
                {
                    SignatureSettings settings = new SignatureSettings()
                    {
                        SignatureLineId = item.Key,
                        SignatureLineImage = File.ReadAllBytes(item.Value),
                        Comments = "Approved",
                        SignTime = DateTime.Now
                    };

                    document.AddDigitalSignature(certificate, settings);
                }

                // Saves the signed document.
                document.Save(Path.GetFullPath(@"Output\Result.docx"), FormatType.Docx);
            }
        }
    }
}

