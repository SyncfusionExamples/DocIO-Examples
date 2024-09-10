using Syncfusion.DocIO.DLS;
using Syncfusion.DocIO;
using Syncfusion.DocIORenderer;

//Get the Source document names from the folder.
string[] sourceDocumentNames = Directory.GetFiles(Path.GetFullPath(@"Data/"));

//Create an WordDocumentinstance for destination document.
using (WordDocument destinationDocument = new WordDocument())
{
    //Add an section and paragraph to the destination document.
    IWSection section = destinationDocument.AddSection();
    IWParagraph paragraph = section.AddParagraph();
    //Appends the TOC field with LowerHeadingLevel and UpperHeadingLevel to determines the TOC entries.
    paragraph.AppendTOC(1, 3);

    //Merge each source document to the destination document.
    foreach (string subDocumentName in sourceDocumentNames)
    {
        //Open the source document files as a stream.
        using (FileStream sourceDocumentPathStream = new FileStream(Path.GetFullPath(subDocumentName), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
        {
            //Open the source documents.
            using (WordDocument sourceDocument = new WordDocument(sourceDocumentPathStream, FormatType.Docx))
            {
                //Sets the break-code of First section of source document as NoBreak to avoid imported from a new page
                sourceDocument.Sections[0].BreakCode = SectionBreakCode.NoBreak;
                //Imports the contents of source document at the end of destination document
                destinationDocument.ImportContent(sourceDocument, ImportOptions.UseDestinationStyles);
            }
        }
    }

    //Updates the table of contents
    destinationDocument.UpdateTableOfContents();
    //Save the destination document.
    using (FileStream outputStream = new FileStream(Path.GetFullPath(@"Output/Output.docx"), FileMode.Create, FileAccess.Write))
    {
        destinationDocument.Save(outputStream, FormatType.Docx);
    }
}