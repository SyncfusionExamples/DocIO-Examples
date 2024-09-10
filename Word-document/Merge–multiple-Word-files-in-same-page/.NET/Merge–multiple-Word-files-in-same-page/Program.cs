using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.Runtime.Serialization;

//Get the list of source document to be imported
List<string> sourceFileNames = new List<string>();
sourceFileNames.Add("Data/Addressblock.docx");
sourceFileNames.Add("Data/Salutation.docx");
sourceFileNames.Add("Data/Greetings.docx");

string destinationFileName = "Data/Title.docx";
using (FileStream destinationStreamPath = new FileStream(destinationFileName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
{
    //Opens the destination document
    using (WordDocument destinationDocument = new WordDocument(destinationStreamPath, FormatType.Automatic))
    {
        ImportOtherDocuments(sourceFileNames, destinationDocument);
        //Saves and closes the destination document
        using (FileStream outputStream = new FileStream("Output/Output.docx", FileMode.Create, FileAccess.Write))
        {
            destinationDocument.Save(outputStream, FormatType.Docx);
            destinationDocument.Close();
        }
    }
}

void ImportOtherDocuments(List<string> sourceFiles, WordDocument destinationDocument)
{
    //Iterate through each source document from the list
    foreach (string sourceFileName in sourceFiles)
    {
        //Open source document
        using (FileStream sourceStreamPath = new FileStream(sourceFileName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
        {
            using (WordDocument document = new WordDocument(sourceStreamPath, FormatType.Automatic))
            {
                //Sets the break-code of First section of source document as NoBreak to avoid imported from a new page
                document.LastSection.BreakCode = SectionBreakCode.NoBreak;
                //Imports the contents of source document at the end of destination document
                destinationDocument.ImportContent(document, ImportOptions.UseDestinationStyles);
                //Close the document.
                document.Close();
            }
        }
    }
}