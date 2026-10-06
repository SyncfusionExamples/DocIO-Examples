# Validate a digital signature in a Word document using C#

The [.NET Word Library](https://www.syncfusion.com/document-sdk/net-word-library) (DocIO) empowers you to create, read, and edit Word documents programmatically without Microsoft Word or interop dependencies. Using this library, you can **validate digital signatures in a Word document** using C#.

## Steps to validate digital signatures in a Word document programmatically

Step 1: Create a new .NET Core console application project.

Step 2: Install the [Syncfusion.DocIO.Net.Core](https://www.nuget.org/packages/Syncfusion.DocIO.Net.Core) NuGet package as a reference to your project from [NuGet.org](https://www.nuget.org/).

Step 3: Include the following namespaces in the Program.cs file.

```csharp
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.Office;
```

Step 4: Add the following code snippet in Program.cs file to validate digital signatures in a Word document.

```csharp
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
```

More information about validate a digital signature in a Word document can be refer in this [documentation](https://help.syncfusion.com/document-processing/word/word-library/net/working-with-digital-signature) section.
