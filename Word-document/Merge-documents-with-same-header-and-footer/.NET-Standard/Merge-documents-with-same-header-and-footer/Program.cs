using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.Office;
using System;
using System.Collections.Generic;
using System.IO;

namespace Merge_documents_with_same_header_and_footer
{
    class Program
    {
        static void Main(string[] args)
        {
            Syncfusion.Licensing.SyncfusionLicenseProvider.RegisterLicense("Mgo+DSMBMAY9C3t2VVhkQlFadV5JXGFWfVJpTGpQdk5xdV9DaVZUTWY/P1ZhSXxRd0djXn5ZcXVQRWVfVEA=");
            //Open the destination document as a stream.
            using (FileStream destinationStreamPath = new FileStream(Path.GetFullPath(@"../../../Data/DestinationDocument.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                //Open the destination document.
                using (WordDocument destinationDocument = new WordDocument(destinationStreamPath, FormatType.Automatic))
                {
                    //Get the Source document names from the folder.
                    string[] sourceDocumentNames = Directory.GetFiles(@"../../../Data/SourceDocuments/");
                    //Merge each source document to the destination document.
                    foreach (string subDocumentName in sourceDocumentNames)
                    {
                        //Open the source document files as a stream.
                        using (FileStream sourceDocumentPathStream = new FileStream(Path.GetFullPath(subDocumentName), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                        {
                            //Open the source documents.
                            using (WordDocument sourceDocuments = new WordDocument(sourceDocumentPathStream, FormatType.Docx))
                            {
                                //Iterate source document sections.
                                foreach (WSection sourceDocumentSections in sourceDocuments.Sections)
                                {
                                    //Clear the headers and footers of the source documents.
                                    sourceDocumentSections.HeadersFooters.FirstPageHeader.ChildEntities.Clear();
                                    sourceDocumentSections.HeadersFooters.FirstPageFooter.ChildEntities.Clear();
                                    sourceDocumentSections.HeadersFooters.EvenHeader.ChildEntities.Clear();
                                    sourceDocumentSections.HeadersFooters.EvenFooter.ChildEntities.Clear();
                                    sourceDocumentSections.HeadersFooters.OddHeader.ChildEntities.Clear();
                                    sourceDocumentSections.HeadersFooters.OddFooter.ChildEntities.Clear();
                                }
                                //Import source documents to destination document.
                                destinationDocument.ImportContent(sourceDocuments);                               
                            }
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
