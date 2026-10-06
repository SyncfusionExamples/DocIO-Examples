# Add a digital signature in a Word document using C#

The [.NET Word Library](https://www.syncfusion.com/document-sdk/net-word-library) (DocIO) empowers you to create, read, and edit Word documents programmatically without Microsoft Word or interop dependencies. Using this library, you can **add an invisible digital signature to a Word document** using C#.

## Steps to add an invisible digital signature to a Word document programmatically

Step 1: Create a new .NET Core console application project.

Step 2: Install the [Syncfusion.DocIO.Net.Core](https://www.nuget.org/packages/Syncfusion.DocIO.Net.Core) NuGet package as a reference to your project from [NuGet.org](https://www.nuget.org/).

Step 3: Include the following namespaces in the Program.cs file.

```csharp
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.Office;
```

Step 4: Add the following code snippet in Program.cs file to add an invisible digital signature to a Word document.

```csharp
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
```

More information about add a digital signature in a Word document can be refer in this [documentation](https://help.syncfusion.com/document-processing/word/word-library/net/working-with-digital-signature) section.
