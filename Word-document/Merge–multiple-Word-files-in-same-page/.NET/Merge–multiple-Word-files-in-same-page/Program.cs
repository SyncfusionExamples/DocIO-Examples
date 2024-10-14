using Syncfusion.DocIO; 
using Syncfusion.DocIO.DLS;

//Register Syncfusion license
Syncfusion.Licensing.SyncfusionLicenseProvider.RegisterLicense("Mgo+DSMBMAY9C3t2UlhhQlNHfV5DQmBWfFN0QXNYfVRwdF9GYEwgOX1dQl9nSXZTc0VlWndfcXNSQWc=");

//Get the list of source documents to be imported.
List<string> sourceFileNames = new List<string>(); 
sourceFileNames.Add("Data/Addressblock.docx"); 
sourceFileNames.Add("Data/Salutation.docx"); 
sourceFileNames.Add("Data/Greetings.docx");

//Get the full path of the destination document.
string destinationFileName = Path.GetFullPath(@"Data/Title.docx");
//Open a file stream to the destination document.
using (FileStream destinationStreamPath = new FileStream(destinationFileName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)) 
{
    //Opens the destination document.
    using (WordDocument destinationDocument = new WordDocument(destinationStreamPath, FormatType.Automatic)) 
    {
        //Calls the method to import other documents into the destination document.
        ImportOtherDocuments(sourceFileNames, destinationDocument); 
        //Saves and closes the destination document.
        using (FileStream outputStream = new FileStream(Path.GetFullPath(@"Output/Output.docx"), FileMode.Create, FileAccess.Write)) 
        {
            destinationDocument.Save(outputStream, FormatType.Docx); 
            //Closes the destination document.
            destinationDocument.Close(); 
        }
    }
}

/// <summary>
/// Import content from multiple source Word documents into a destination document.
/// </summary>
void ImportOtherDocuments(List<string> sourceFiles, WordDocument destinationDocument) //Method to import content from source documents into the destination document
{
    //Iterate through each source document from the list.
    foreach (string sourceFileName in sourceFiles) 
    {
        //Open the source document.
        using (FileStream sourceStreamPath = new FileStream(sourceFileName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)) 
        {
            //Opens the Word document from the file stream.
            using (WordDocument document = new WordDocument(sourceStreamPath, FormatType.Automatic)) //Initialize a WordDocument object for the source document
            {
                //Set the break-code of the first section of the source document as NoBreak to avoid starting content on a new page.
                document.LastSection.BreakCode = SectionBreakCode.NoBreak; 
                //Imports the content of the source document at the end of the destination document.
                destinationDocument.ImportContent(document, ImportOptions.UseDestinationStyles); 
                //Close the source document
                document.Close(); 
            }
        }
    }
}
