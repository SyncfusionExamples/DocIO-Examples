using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;

//Open the source document
using (WordDocument sourceDocument = new WordDocument(@"../../../Data/SourceDocument.docx"))
{

    //Open the destination document
    using (WordDocument destinationDocument = new WordDocument(@"../../../Data/DestinationDocument.docx"))
    {

        //Import the contents; the first source section is merged into the last destination section
        destinationDocument.ImportContent(sourceDocument, ImportOptions.UseDestinationStyles, true);

        //Save the destination document
        destinationDocument.Save("../../../Output/CombinedDocument.docx", FormatType.Docx);
    }
}