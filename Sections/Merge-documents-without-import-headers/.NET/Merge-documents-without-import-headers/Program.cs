using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;

namespace Merge_documents_without_import_headers
{
    class Program
    {
        static void Main(string[] args)
        {
            //Load the destination Word document as a stream.
            using (FileStream destinationStreamPath = new FileStream(Path.GetFullPath(@"Data/DestinationDocument.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                //Open the destination Word document.
                using (WordDocument destinationDocument = new WordDocument(destinationStreamPath, FormatType.Automatic))
                {
                    //Load the source Word document as a stream.
                    using (FileStream sourceDocumentPathStream = new FileStream(Path.GetFullPath(@"Data/SourceDocument.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                    {
                        //Open the source Word document.
                        using (WordDocument sourceDocument = new WordDocument(sourceDocumentPathStream, FormatType.Docx))
                        {
                            //Iterate source Word document sections.
                            foreach (WSection sourceDocumentSection in sourceDocument.Sections)
                            {
                                //Remove the first page header.
                                sourceDocumentSection.HeadersFooters.FirstPageHeader.ChildEntities.Clear();
                                //Remove the first page footer.
                                sourceDocumentSection.HeadersFooters.FirstPageFooter.ChildEntities.Clear();
                                //Remove the even header.
                                sourceDocumentSection.HeadersFooters.EvenHeader.ChildEntities.Clear();
                                //Remove the even footer.
                                sourceDocumentSection.HeadersFooters.EvenFooter.ChildEntities.Clear();
                                //Remove the odd header.
                                sourceDocumentSection.HeadersFooters.OddHeader.ChildEntities.Clear();
                                //Remove the odd footer.
                                sourceDocumentSection.HeadersFooters.OddFooter.ChildEntities.Clear();
                            }
                            //Merge source Word document content to destination Word document.
                            destinationDocument.ImportContent(sourceDocument);
                        }
                    }
                    //Create a file stream.
                    using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"Output/Sample.docx"), FileMode.Create, FileAccess.ReadWrite))
                    {
                        //Save the Word document to the file stream.
                        destinationDocument.Save(outputFileStream, FormatType.Docx);
                    }
                }
            }
        }
    }
}
