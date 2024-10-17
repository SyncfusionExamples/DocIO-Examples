using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;

//Create a list and add the paths of the source Word documents to it.
List<string> sourceFileNames = new List<string>();
sourceFileNames.Add("Data/Addressblock.docx");
sourceFileNames.Add("Data/Salutation.docx");
sourceFileNames.Add("Data/Greetings.docx");

//Get the absolute path of the destination Word document.
string destinationFileName = Path.GetFullPath(@"Data/Title.docx");
using (FileStream destinationStream = new FileStream(destinationFileName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
{
    //Open the destination document.
    using (WordDocument destinationDocument = new WordDocument(destinationStream, FormatType.Automatic))
    {
        ImportOtherDocuments(sourceFileNames, destinationDocument);
        //Save the destination document.
        using (FileStream outputStream = new FileStream(Path.GetFullPath(@"Output/Output.docx"), FileMode.Create, FileAccess.Write))
        {
            destinationDocument.Save(outputStream, FormatType.Docx);
        }
    }
}

/// <summary>
/// Import content from multiple source Word documents into a destination document.
/// </summary>
void ImportOtherDocuments(List<string> sourceFiles, WordDocument destinationDocument)
{
    //Iterate through each source document from the list.
    foreach (string sourceFileName in sourceFiles)
    {
        using (FileStream sourceStream = new FileStream(sourceFileName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
        {
            //Open the source document.
            using (WordDocument sourceDocument = new WordDocument(sourceStream, FormatType.Automatic))
            {
                //Set the break-code of First section of source document as NoBreak to avoid imported from a new page.
                sourceDocument.LastSection.BreakCode = SectionBreakCode.NoBreak;
                //Import the contents of source document at the end of destination document.
                destinationDocument.ImportContent(sourceDocument, ImportOptions.UseDestinationStyles);
            }
        }
    }
}