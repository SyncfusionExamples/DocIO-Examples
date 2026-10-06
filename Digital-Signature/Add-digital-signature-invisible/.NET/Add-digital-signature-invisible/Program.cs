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
            //Sets the application version used to create the signature.
            settings.ApplicationVersion = "16.0";
            //Sets the Office version recorded with the signature.
            settings.OfficeVersion = "16.0";
            //Sets the Windows version recorded with the signature.
            settings.WindowsVersion = "10.0";
            //Sets the horizontal resolution recorded with the signature.
            settings.HorizontalResolution = 1920;
            //Sets the vertical resolution recorded with the signature.
            settings.VerticalResolution = 1080;
            //Sets the color depth recorded with the signature.
            settings.ColorDepth = 32;
            //Sets the cryptographic provider identifier.
            settings.ProviderId = new Guid("00000000-0000-0000-0000-000000000000");
            //Adds an invisible digital signature to the document using the certificate and settings.
            document.AddDigitalSignature(certificate, settings);
            //Saves the Word document to file.
            document.Save(Path.GetFullPath(@"Output\Result.docx"), FormatType.Docx);
            //Closes the document
            document.Close();
        }
    }
}
