using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.Office;

namespace Add_digital_signature_invisible
{
    class Program
    {
        static void Main(string[] args)
        {
            //Opens an existing Word document.
            WordDocument document = new WordDocument(Path.GetFullPath(@"Data\Template.docx"));
            //Loads the signing certificate from disk.
            OfficeDigitalSignatureCertificate certificate = new OfficeDigitalSignatureCertificate(Path.GetFullPath(@"Data\certificate.pfx"), "password123");
            //Configures signature settings.
            SignatureSettings settings = new SignatureSettings();
            settings.Comments = "Approved";
            settings.SignTime = DateTime.Now;
            //Adds an invisible digital signature to the document using the certificate and settings.
            document.AddDigitalSignature(certificate, settings);
            //Saves the Word document to file.
            document.Save(Path.GetFullPath(@"Output\Result.docx"), FormatType.Docx);
            //Closes the document
            document.Close();
        }
    }
}
