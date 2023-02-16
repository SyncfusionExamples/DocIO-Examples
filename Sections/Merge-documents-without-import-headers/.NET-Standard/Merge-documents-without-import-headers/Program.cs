using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;

namespace Merge_documents_without_import_headers
{
    class Program
    {
        static void Main(string[] args)
        {
            Syncfusion.Licensing.SyncfusionLicenseProvider.RegisterLicense("Mgo+DSMBMAY9C3t2VVhkQlFadV5JXGFWfVJpTGpQdk5xdV9DaVZUTWY/P1ZhSXxRd0djXn5ZcXVQRWVfVEA=");
            //Open the file as a stream.
            using (FileStream destinationStreamPath = new FileStream(Path.GetFullPath(@"../../../DestinationDocument.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                //Load the file stream into a Word document.
                using (WordDocument destinationDocument = new WordDocument(destinationStreamPath, FormatType.Automatic))
                {
                    //Open the source document as a stream.
                    using (FileStream sourceDocumentPathStream = new FileStream(Path.GetFullPath(@"../../../SourceDocument.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                    {
                        //Open the source document.
                        using (WordDocument sourceDocument = new WordDocument(sourceDocumentPathStream, FormatType.Docx))
                        {
                            //Iterate source document.
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
                            destinationDocument.ImportContent(sourceDocument);
                        }
                    }
                    //Create a file stream.
                    using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"../../../Sample.docx"), FileMode.Create, FileAccess.ReadWrite))
                    {
                        //Save the Word document to the file stream.
                        destinationDocument.Save(outputFileStream, FormatType.Docx);
                    }
                }
            }
        }
    }
}
