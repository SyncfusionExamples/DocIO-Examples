using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;

//Opens the source document
FileStream SourceFileStream = new FileStream(@"../../../Data/SourceDocument.docx", FileMode.Open);
WordDocument sourceDocument = new WordDocument(SourceFileStream, FormatType.Docx);
//Opens the destination document  
FileStream DestinationFileStream = new FileStream(@"../../../Data/DestinationDocument.docx", FileMode.Open);
WordDocument destinationDocument = new WordDocument(DestinationFileStream, FormatType.Docx);
//Processes the body contents for each section in the Word document
foreach (WSection section in sourceDocument.Sections)
{
    //Accesses the body of section where all the contents in document are apart
    foreach (TextBodyItem bodyItem in section.Body.ChildEntities)
    {
        destinationDocument.LastSection.Body.ChildEntities.Add(bodyItem.Clone());
    }
}
//Closes the source document
sourceDocument.Close();
destinationDocument.Save(@"../../../Output/Output.docx", FormatType.Docx);
//Closes the destination document
destinationDocument.Close();