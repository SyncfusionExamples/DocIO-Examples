using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.Office;

namespace Sign_signature_line
{
    class Program
    {
        static void Main(string[] args)
        {
            //Opens an existing Word document with signature line.
            using (WordDocument document =
                new WordDocument(
                    Path.GetFullPath(@"Data\Template.docx")))
            {
                WPicture picture = null;
                OfficeSignatureLine signatureLine = null;

                //Finds the signature line in the document.
                foreach (WSection section in document.Sections)
                {
                    foreach (WParagraph paragraph in section.Paragraphs)
                    {
                        foreach (Entity entity in paragraph.ChildEntities)
                        {
                            if (entity is WPicture currentPicture && currentPicture.IsSignatureLine)
                            {
                                picture = currentPicture;
                                signatureLine = picture.SignatureLine;
                                break;
                            }
                        }

                        if (picture != null)
                            break;
                    }

                    if (picture != null)
                        break;
                }

                //Loads the signing certificate.
                OfficeDigitalSignatureCertificate certificate =
                    new OfficeDigitalSignatureCertificate(
                        Path.GetFullPath(@"Data\certificate.pfx"),
                        "password123");

                //Signs the signature line with the image.
                SignatureSettings settings = new SignatureSettings
                {
                    SignatureLineId = signatureLine.Id,
                    SignTime = DateTime.Now,
                    SignatureLineImage = File.ReadAllBytes(Path.GetFullPath(@"Data\Signature.png")),
                    Comments = "Approved"
                };

                if (File.Exists(Path.GetFullPath(@"Data\Signature.png")))
                    settings.SignatureLineImage = File.ReadAllBytes(Path.GetFullPath(@"Data\Signature.png"));

                document.AddDigitalSignature(certificate, settings);
                //Saves the signed document.
                document.Save(Path.GetFullPath(@"Output\Result.docx"), FormatType.Docx);
            }

        }
    }
}
