using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.Office;

namespace Sign_signature_line
{
    class Program
    {
        static void Main(string[] args)
        {
            //Opens an existing Word document and adds a signature line.
            Guid signatureLineId;
            using (WordDocument document = new WordDocument(Path.GetFullPath(@"Data\Template.docx")))
            {
                //Gets the first section of the document.
                IWSection section = document.Sections[0];
                //Adds new paragraph to the section.
                IWParagraph paragraph = section.AddParagraph();
                //Adds new text to the paragraph
                paragraph.AppendText("Signed by: ");
                //Adds a new paragraph that will host the signature line and aligns it to the left.
                IWParagraph signatureParagraph = section.AddParagraph();
                signatureParagraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Left;
                //Configures the signature line settings.
                SignatureLineSettings lineSettings = new SignatureLineSettings();
                lineSettings.Signer = "John Doe";
                lineSettings.SignerTitle = "Manager";
                lineSettings.Email = "john.doe@example.com";
                lineSettings.Instructions = "Please review and sign.";
                lineSettings.ShowDate = true;
                //Inserts the signature line into the new paragraph with the specified dimensions.
                IWPicture picture = signatureParagraph.AppendSignatureLine(lineSettings, 200, 100);
                //Gets the unique identifier of the inserted signature line.
                signatureLineId = ((WPicture)picture).SignatureLine.Id;
                //Saves the Word document that contains the signature line.
                document.Save(Path.GetFullPath(@"Output\Result.docx"), FormatType.Docx);
            }

            //Loads the signing certificate from disk.
            OfficeDigitalSignatureCertificate certificate = new OfficeDigitalSignatureCertificate(Path.GetFullPath(@"Data\certificate.pfx"), "password123");
            //Opens the saved document to sign the signature line.
            using (WordDocument document = new WordDocument(Path.GetFullPath(@"Output\Result.docx")))
            {
                //Configures signature settings, binds the signature to the signature line, and supplies a custom signature image.
                SignatureSettings settings = new SignatureSettings();
                settings.SignatureLineId = signatureLineId;
                settings.SignatureLineImage = File.ReadAllBytes(Path.GetFullPath(@"Data\Signature.png"));
                settings.Comments = "Approved";
                settings.SignTime = DateTime.Now;
                //Signs the signature line with the supplied image to apply a visible digital signature.
                document.AddDigitalSignature(certificate, settings);
                //Saves the signed Word document to file.
                document.Save(Path.GetFullPath(@"Output\Result.docx"), FormatType.Docx);
            }
        }
    }
}
